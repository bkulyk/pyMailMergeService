#!/usr/bin/env python
#Copyright (c) 2010, Brian Kulyk
#All rights reserved.
#
#Redistribution and use in source and binary forms, with or without
#modification, are permitted provided that the following conditions are met:
#1. Redistributions of source code must retain the above copyright
#   notice, this list of conditions and the following disclaimer.
#2. Redistributions in binary form must reproduce the above copyright
#   notice, this list of conditions and the following disclaimer in the
#   documentation and/or other materials provided with the distribution.
#3. All advertising materials mentioning features or use of this software
#   must display the following acknowledgement:
#   This product includes software developed by the <organization>.
#4. Neither the name of the Brian Kulyk nor the
#   names of its contributors may be used to endorse or promote products
#   derived from this software without specific prior written permission.
#
#THIS SOFTWARE IS PROVIDED BY Brian Kulyk ''AS IS'' AND ANY
#EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
#WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
#DISCLAIMED. IN NO EVENT SHALL Brian Kulyk BE LIABLE FOR ANY
#DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
#(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
#LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
#ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
#(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
#SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""
@todo:  add support for batches of documents, ie a real mail merge
@todo:  for the batch support, probably need to be able to support receiving merge values from XML or JSON maybe
@todo:  turn into a package, add a setup.py http://www.packtpub.com/article/writing-a-package-in-python
@todo:  Before writing final xml files to the new odt. Make sure there are not any macros. Marcos could cause harm to the system.
@todo:  Possible improve the doc/docx support, it seems to work good, and fast, but it 
        would probably be better if it wasn't converting to .odt everytime, and just cached
        the converted file.
"""
from sys import path
#path.append( '/usr/lib/openoffice.org/program/' ) #for document converter
path.append( "/usr/share/pyshared/" ) #for uno library used by DocumentConverter
import base64                       #to be able to return the output file utf encoded
import zipfile                      #to open odt so we can do find and replace on the xml
import re                           #regular expressions for doing find and replace on the xml
import os                           #removing the temp files
import tempfile                     #create temp files for writing xml and output
from DocumentConverter import *     #to convert open office documents
from OpenOfficeDocument import OpenOfficeDocument
from SOAPpy import *                #soap interface
from lxml import etree              #parsing xml
import codecs                       #opening a file for writing as UTF-8
import logging                      #for logging
from logging import handlers        #to be able to do the rotating log, with formatting
class pyMailMergeService:
        appName = 'pyMailMergeService'
        enablelogging = True
        logging = None
        #NOTE: this is not all of the namespaces that are in the document, just the ones I need, plus a few.
        ns = {'table':"urn:oasis:names:tc:opendocument:xmlns:table:1.0", 
              'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0' ,
              'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
              'style':'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
              'draw':'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
              'xlink':'http://www.w3.org/1999/xlink' }
        xml = None
        convertsionmap = { 'doc':'odt', 'docx':'odt', 'rtd':'odt', 'xls':'ods', 'xlsx':'ods' }
        #for caching regular expressions
        _regExString = { 'amp':r"&(?!amp;|#[0-9]{2,5};|[a-z]{2,5};)",
                         'tokens':r'~([a-zA-Z\|]+::\w+[\w\|\)]*)~',
                         'repeatrow':r"^repeatrow|",
                         'repeatcol':r"^repeatcolumn|",
                         'htmlparagraph':'\<p\>.+\<\/p\>'
                         }
        _regExCompiled = {}
        def __init__( self, **kwargs ):
            """enable or disable logging"""
            #self.soapContext = kwargs.get( 'soapContext', None )
            self.enablelogging = kwargs.get( 'enablelogging', True )
            if self.enablelogging:
                self.logging = logging.getLogger( self.appName )
                self.logging.setLevel( logging.DEBUG )
                handler = logging.handlers.RotatingFileHandler( 'pymms.log' )
                formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
                handler.setFormatter( formatter )
                self.logging.addHandler( handler )
                self.logging.info( "Starting %s listing on %s:%s" % ( self.appName, kwargs.get( 'host', '?' ), kwargs.get( 'port', '?' ) ) )
            else:
                self.logging = loggingVoid()
        def _getRegEx( self, regex ):
            """
            This is here an an effort to stop compiling the same regular expressions over and 
            over again.  Multiple times per page hit.  Just doesn't seem very efficient.
            This also turned out to be extremely valuable for unit testing, as it savings having to 
            re-type/change regular expressions when testing.
            """
            if regex in self._regExCompiled.keys():
                return self._regExCompiled[ regex ]
            else:
                self._regExCompiled[ regex ] = re.compile( self._regExString[ regex ] )
                return self._regExCompiled[ regex ]
        def getMethods(self):
            x = []
            for i in dir( self ):
                if i[0] != '_' and i != 'getMethods':
                    if callable( getattr( self, i ) ):
                        x.append( i )
                        self.logging.info( 'Exposing method for webservice consumption: "%s"' % i )
            return x
            #return self.soapContext
        def _logSoapInfo( self, _SOAPContext=None ):
            """This is a method so we can start logging some information about the reqests being made.
            This logs the name of the method being called, and the"""
            if _SOAPContext is not None:
                headers = "".join(_SOAPContext.httpheaders.headers ).replace( "\n", ' ' )
                self.logging.info( "Soap Request - %s - User-agent %s" % ( _SOAPContext.soapaction, _SOAPContext.httpheaders.dict['user-agent'] ) )
        def hello( self, param0='world', _SOAPContext=None ):
            self._logSoapInfo( _SOAPContext )
            """This is a simple 'Hello World' function to test soap calls"""
            self.logging.info( 'hello method called from value: %s' % ( param0 ) )
            return u"hello %s" % param0
        def getTokens( self, param0, _SOAPContext=None ):
            """Parse the odt for tokens that should be replaced with content, kind 
            of like a mailmerge.  Tokens are in the format of ~token::name~"""
            self._logSoapInfo( _SOAPContext )
            if not os.path.exists( os.path.abspath( "docs/%s" % param0 ) ):
                self.logging.error( "templatefile: %s does not exist" % os.path.abspath( "docs/%s" % param0 ) )
                return "error: templatefile: %s does not exist" % param0
            self.logging.info( "getting tokens for document: %s" % param0 )
            odtName, zip = self._get_source( param0 )
            #get the tokens from the xml
            exp = self._getRegEx( 'tokens' )
            matches = []
            #I found that I need to look for tokens in the styles and meta fils as well, because meta has the document title, and styles has the content for the document footers and probably headers
            for file in [ u'content.xml', u'styles.xml', u'meta.xml' ]:
                xml = zip.read( file )
                matches.extend( exp.findall( xml ) )
                #collect a list of all the images in the document
            self._close_source( param0, odtName, zip )
            matches = list( set( matches ) ) #removes duplicate items from the list
            self.logging.info( "tokens sent: %s" % len( matches ) )
            return matches
        def _get_source( self, param0 ):
            """get the source document, (which may include converting a doc/docx file to odt) 
            open the zip and return the filename and zip"""
            #http://docs.python.org/library/zipfile.html
            if self._getFileExtension( param0 ) == 'doc':
                odtName = self._getTempFile( ".odt" )
                self.converter.convert( param0, odtName )
            else:
                odtName = param0
            zip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            return odtName, zip
        def _close_source(self, param0, odtName, zip ):
            """
            close the zip file and delete the temprary odt if converted a doc/docx to odt
            """
            zip.close()
            #remove the temporary doc -> odt conversion file
            if self._getFileExtension( param0 ) == 'doc':
                os.unlink( odtName )
        def pdf( self, param0, param1, _SOAPContext=None ):
            """Convert the odt provided in param0 to a pdf, using the values from param1 
            as mailmerge values."""
            self._logSoapInfo( _SOAPContext )
            return self.convert( param0, param1, 'pdf', None )
        def convert( self, param0, param1, param2, _SOAPContext=None ):
            """
            open the odt file which is really just a zip, create a new odt (zip), and copy all
            files from the first odt to the new.  If the file is the content.xml, then 
            replace the tokens from params in the xml before placing it in the new odt.
            Then create a PDF document out of the new odt, and return it base 64 encoded, so 
            it can be transported via Soap.  Soap doesn't seem to like binary data.
            """
            self._logSoapInfo( _SOAPContext )
            self.logging.info( "converting %s docuemnt to %s" % (param0, param2) )
            #if not os.path.exists( param0 ):
            #    return "error: templatefile: %s does not exist" % param0
            if self._getFileExtension( param0 ) == 'doc':
                odtName = self._getTempFile( ".odt" )
                #self.converter.convert( param0, odtName )
                OpenOfficeDocument.convert( param0, odtName )
            else:
                odtName = param0
            '''this has been moved here beause it's more effiecent then running multiple times, and does 
            not cause caching problems accross multiple soap requests.  Side note: Man I love unit tests'''
            skipMerge = False
            try:
                params = param1['item'] #no idea why this stuff is inside of 'item'...
                params = self._sortparams( params )
            except:
                """if there is only a single value passed it will come out as a struct instead of a list"""
                try:
                    params = [ {} ]
                    for k,x in param0:
                        params[0] = dict( )
                        params[0][ k ] = []
                        for y in x:
                            params[0][k].append( y )
                except:
                    #cant do merge, bad data given
                    return None
            doctype = param2
            '''http://docs.python.org/library/shutil.html'''
            sourcezip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            #create destination zip
            destodt = self._getTempFile( '.odt' )
            destzip = zipfile.ZipFile( destodt, 'w' )
            #copy all file from the source zip(odt) the the destination zip
            tmpfiles = {}
            xmlFileList = ['content.xml', 'meta.xml', 'styles.xml']
            """Write all of the contents of the sourcezip to the destzip if the the filename is not 
            in the xmlFileList list"""
            for x in sourcezip.namelist():
                if x not in xmlFileList:
                    tmp = sourcezip.read( x )
                    destzip.writestr( x, tmp )
            """these files need to be done last, in case there is any modifiers that want to 
            change the contents of the zip, for example the image| modifier""" 
            for x in xmlFileList: 
                tmp = sourcezip.read( x )
                tmp = self._replaceContent( sourcezip.read( x ), params, destzip )
                '''I could not get the contents of the xml (tmp) file to write directly
                to the zip because of file encoding stuff with french characters, so I'm writing
                the contents to a utf-8 file, then putting that file into the zip archive.'''
                tmpFileName = "%s" % self._getTempFile()
                tmpFile = codecs.open( tmpFileName, 'w', 'utf-8' )
                try:
                    tmpFile.write( unicode( tmp, 'utf-8' ) ) #for english
                except:
                    tmpFile.write( tmp ) #for french
                tmpFile.close()
                destzip.write( tmpFileName, x )
                #remove file now
                os.unlink( tmpFileName )
            #clean up
            destzip.close()
            sourcezip.close()
            #now convert the odt to pdf
            destpdf = u"%s" % ( self._getTempFile( "."+doctype ) )
            unlinkODT = True
            try:
                #this is now going to create an instance of converter on demand in case OpenOffice
                #crashes and needs to restart.
#                converter = DocumentConverter()
#                converter.convert( destodt, destpdf )
                OpenOfficeDocument.convert( destodt, destpdf )
            except:
                unlinkODT = False
                errormsg = "Could not convert document, usually bad xml.  Check ODT File: '%s'" % destodt
                self.logging.error( errormsg )
                return "error: %s" % errormsg
            f = open( destpdf, 'r' )
            doc = f.read()
            f.close()
            if unlinkODT:
                os.unlink( destodt ) #delete the odt file
            os.unlink( destpdf ) #delete the pdf file
            #remove the temporary doc -> odt conversion file
            if self._getFileExtension( param0 ) == 'doc':
                os.unlink( odtName )
            self.logging.info( "done converting document: %s" % odtName )
            return base64.b64encode( doc ) #pdf
        def _replaceContent( self, xml, params, zip ):
            """
            Do a regular expression find and replace for all of the params, unless the 
            param has multiple values, then do the duplicate rows stuff.
            """
            #these were moved out of the loop because compiling the re everytime would be ineffeciant
            para_exp = re.compile( '\<p\>.+\<\/p\>' )
            #for syntax: http://www.php2python.com/wiki/control-structures.foreach/
            for dictionary in params:
                key = dictionary.keys()[0]
                value = dictionary.values()[0]
                if type( value ).__name__ in ( 'instance', 'list', 'typedArrayType' ):
                    xml = self._multipleValues( xml, key, value )
                else:
                    if key.find( r'image|' ) == 0:
                        self._image( key, value, xml, zip )
                    if key.find( r'multiparagraph|' ) == 0:
                        xml = self._multiparagraph( key, value, xml )
                    if key.find( r"if|" ) == 0:
                        xml = self._if( key, value, xml )
                    if key.find( r"endif|" ) == 0:
                        continue
                    if key.find( r"repeatsection|" ) == 0:
                        xml = self._repeatsection( key, value, xml )
                    if key.find( r"endrepeatsection|" ) == 0:
                        continue
                    #do a simple find and replace with a regular expression
                    exp = re.compile( '~%s~' % re.escape(key) )
                    value = self._cleanValue( value )
                    try:
                        xml = exp.sub( "%s" % value, xml )
                    except:
                        #This took me hours and hours to figure out.  Apparantly sometimes when
                        #you copy and paste french content into an odt, it does not save as proper
                        #utf-8, so it needs a conversion before we can continue.
                        xml = unicode( xml, "utf-8" )
                        xml = xml.replace( "~%s~" % key, value )
            return xml
        def _image( self, key, value, xml, zip ):
            """This modifier is for replacing the contents of an image. A good example for 
            this would be a product where you are creating invoices on behalf of multiple 
            companies and you need to be able to put a specific company's logo at the
            top of the document without having to change the document it's self.  The
            value should be a string containin the contents of the image base64 encoded."""
            logging.info( "image modifier found." )
            if value is not None:
                x = self._getXML( xml )
                drawframe = x.xpath( "//draw:frame[@draw:name='~%s~']/draw:image" % key, namespaces=self.ns )
                if len( drawframe ):
                    href = drawframe[0].get( '{%s}href' % self.ns['xlink'] )
                    zip.writestr( href, base64.b64decode( value ) )
            else:
                logging.error( "no image data provdided." )
        def _if( self, key, value, xml ):
            """
            If there is a modifier tag for an if statement, only show the stuff between the if
            and the endif modifiers IF the value is 1, '1' or true.
            This piece is really touchy as the start and end tags have to be in the same 
            depth of xml nodes, which can be difficult.  This whole piece may need be be
            reworked to do the editing via XML instead of text
            """
            self.logging.info( "if modifier found." )
            xmlbackup = xml
            starttoken = "~"+key+"~"
            closetoken = starttoken.replace( r"if|", r"endif|" ) 
            startpos = xml.find( starttoken )
            closepos = xml.find( closetoken )
            if startpos == -1 and closepos == -1:
                return xml
            tmp = None
            if startpos > 0 and closepos > 0 and value == 0:
                tmp = xml[ 0:startpos ] + xml[ closepos + len( closetoken ): ]
            if tmp is not None:
                tmp = tmp.replace( closetoken, '' )
                tmp = tmp.replace( starttoken, '' )
            else:
                xmlbackup = xmlbackup.replace( closetoken, '' )
                xmlbackup = xmlbackup.replace( starttoken, '' )
                self.logging.error( "if modifier failed for token %s, value is 0.  Reverting to old XML as not to completely blow up the document." % key )
                return xmlbackup
            #The removal of content in the manner that it was just done could break the 
            #xml structure and cause it to not parse anymore.  That would be bad.  So if the 
            #xml is broken then don't use this modified version and fall back to the old.
            try:
                xml = etree.XML( tmp )
                return tmp
            except:
                self.logging.error( "if modifier failed for token %s, XML can't parse.  Reverting to old XML as not to completely blow up the document." % key  )
                return xmlbackup
        def _repeatsection( self, key, value, xml ):
            """
            Repeat a section of the document as many times as the value says, which should be 
            and integer.
            """
            self.logging.info( 'repeating section found.' )
            xmlbackup = xml
            starttoken = "~"+key+"~"
            closetoken = starttoken.replace( r"repeatsection|", r"endrepeatsection|" ) 
            startpos = xml.find( starttoken ) + len( starttoken )
            closepos = xml.find( closetoken )
            if startpos == -1 or closepos == -1 or value <= 0:
                return xml
            #now get the content of the section
            content = xml[ startpos : closepos ]
            #duplicate the conent as many times as the value
            allcontent = ""
            for x in range( 1,value ):
                allcontent = allcontent+content
            xml = xml[:closepos] + allcontent + xml[closepos:]
            xml = xml.replace( starttoken, '' )
            xml = xml.replace( closetoken, '' )
            #The removal of content in the manner that it was just done could break the 
            #xml structure and cause it to not parse anymore.  That would be bad.  So if the 
            #xml is broken then don't use this modified version and fall back to the old.
            try:
                tmp = etree.XML( xml ) #if xml is busted this will throw an exception
                return xml
            except:
                self.logging.error( "repeatsection for %s is bad, reverting to backed up xml" % key )
                xmlbackup = xmlbackup.replace( starttoken, '' )
                xmlbackup = xmlbackup.replace( closetoken, '' )
                return xmlbackup
        def _multiparagraph( self, key, value, xml ):
            """
            If there was a single paragraph containing the merge tag, and now there 
            needs to be multiple paragrpahs, convert html <p> tags to their openoffice
            couterparts, which is still a p, but with a namespace and style attribute
            """
            self.logging.info( "multiparagraph modifier found." )
            self.xml = None
            x = self._getXML( xml )
            paragraphs = x.xpath( '//*[contains(text(),"~%s~")]' % key, namespaces=self.ns )
            if not len( paragraphs ):
                return xml
            para = paragraphs[0]
            i=0
            #this is here because I was having a problem where the text was actually 
            #in a span inside the paragraph, and not directly in the paragraph 
            while para.tag != "{%s}p" % self.ns['text'] and i < 5: #5 is arbitrary, should be enough
                para = para.getparent()
                i += 1
            parent = para.getparent()
            #need to add the paragraphy style attribute (and any others) to the new p tags
            attribs = para.attrib
            value = self._cleanValue( value )
            try:
                html = etree.XML( "<html>" + value + "</html>" )
            except:
                html = etree.XML( "<html>" + value + "</html>" )
            previous = para
            for tag in html.findall( 'p' ):
                p = etree.Element( "{%s}p" % self.ns['text'], nsmap=self.ns, attrib=attribs )
                p.text =  tag.text
                #need to add next instead of append because it was messing up ordering of content
                previous.addnext( p )
                previous = p
            parent.remove( para )
            return etree.tostring( x )
        def _multipleValues( self, xml, key, params ):
            """
            There are multiple values for this key, so see if there is a modifier on the key to, 
            repeat columns or rows. If so, then duplicat the column, or row, then fill in the 
            values.  This is like the normal find and replace, but not global.
            """
            rowExp = re.compile( "^" + re.escape( 'repeatrow|' ) )
            colExp = re.compile( "^" + re.escape( 'repeatcolumn|' ) )
            if re.match( rowExp, key ):
                #if it matches the repeat row modifier, then repeat the rows.
                return self._repeatingRow( xml, key, params ) 
            elif re.match( colExp, key ):
                #if it matches the repeat column, t    oken modifer, then repeat the columns.
                return self._repeatcolumn( xml, key, params )
            else:
                self.xml = None #need to clear the xml cache.                
                #if it doesn't match any modifers then, all of the values should already be duplicated, so starting doing find an repalce, but not global
                exp = re.compile( '~%s~' % re.escape(key) )
                for v in params:
                    v = self._cleanValue( v )
                    if key.find( r'multiparagraph|' ) == 0:
                        xml = self._multiparagraph( key, v, xml )
                    else:
                        xml = re.sub( exp, v, xml, count=1 )
                return xml
        def _repeatcolumn( self, xml, key, params ):
            """
            Repeat an entire column of a given table, replacing values as we go.
            """
            self.logging.info( "repeating column found" )
            x = self._getXML( xml )
            cols = x.xpath( '//table:table-row/table:table-cell[contains(.,"%s")]' % key, namespaces=self.ns )
            if len( cols ):
                sourcerow = cols[0].getparent()
                sourceIndex = sourcerow.getparent().index( sourcerow )
                #get the index of the cell
                index = sourcerow.index( cols[0] )
                if len( sourcerow ):
                    table = sourcerow.getparent()
                    i=0
                    for row in table.xpath( "./table:table-row" , namespaces=self.ns ):
                        rowIndex = table.index( row )
                        #need to hold off on repeating the cells in the source column
                        #until last because it would mess up the column count that I need
                        #in order to make the merged cell count stuff work 
                        if int(sourceIndex) != int(rowIndex)+i:
                            insertIndex, fixspanned = self._repeatcolumn_getinsertindex( index, sourcerow, row )
                            cells = row.xpath( "./table:table-cell", namespaces=self.ns )
                            if insertIndex is not None:
                                self._repeatcolumn_dorepeat( row, insertIndex, key, params, dontupdatecolumns=True )
                            if fixspanned is not None:
                                colsspan = cells[ fixspanned ].get( '{%s}number-columns-spanned' % self.ns['table'] )
                                colsspan = int( colsspan ) + len( params ) - 1
                                cells[ fixspanned ].set( '{%s}number-columns-spanned' % self.ns['table'], "%s" % colsspan )
                    #since the source row was skipped, it needs to be done now.
                    self._repeatcolumn_dorepeat( sourcerow, index, key, params )
                    i+=1
                self.xml = etree.tostring( x )
                return self.xml
            else:
                return xml
        def _repeatcolumn_dorepeat( self, row, index, key, params, dontupdatecolumns=False ):
            """
            This method is called by _repeatcolumn.  This method duplicates the cell 
            as many times as indicated by the length of the params.
            """
            cells = row.xpath( "./table:table-cell", namespaces=self.ns )
            if index < 0:
                return
            oldString = etree.tostring( cells[ index ] )
            previous = cells[ index ]
            i=0
            for param in params:
                newElement = oldString.replace( "~%s~" % key, self._cleanValue( param ) )
                newElement = etree.XML( newElement )
                previous.addnext( newElement )
                previous = newElement                
                #this method gets called once for each row, however it only needs to update the columns once per column
                if dontupdatecolumns == False: 
                    #need to updeate th table-column nodes
                    stylename = newElement.get( "{%s}style-name" % self.ns['table'] )
                    stylename, stylenumber = [ stylename[:-1], stylename[-1:] ] #need to extract the number at the end.
                    tablecolumns = row.getparent().xpath( r"./table:table-column[@table:style-name='%s']" % stylename, namespaces=self.ns )
                    if len( tablecolumns ):
                        tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                        tablecolumn = tablecolumns[ len( tablecolumns )-1 ] #should only be one.
                    else: #as a last option, duplicate the first table-column
                        tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                        tablecolumn = tablecolumns[0]
                    tablecolumnString = etree.tostring( tablecolumn )
                    tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                    if len( cells ) < int( len( cells ) + i ):
                        tablecolumn.addnext( etree.XML( tablecolumnString ) )
                i += 1
            row.remove( cells[ index ] )
        def _repeatcolumn_getinsertindex( self, index, row_a, row_b ):
            """
            Because of merged rows and whatnot we may need to do some checking as to which cell
            to actually duplicate.  This method checks what cell to duplicate and retuns the index,
            also if the row has merged cells it returns the index of the column that needs to have
            it's merged cells updated.
            """
            i=0;
            cellsa = row_a.xpath( "./table:table-cell", namespaces=self.ns ) 
            cellsb = row_b.xpath( "./table:table-cell", namespaces=self.ns )
            span_start  = None
            span_size   = None
            span_end    = None
            fixspanned  = None
            span_total  = 0
            if len( cellsa ) == len( cellsb ):
                return [ index, None ]
            else:
                for i in range( len( cellsa ) ):
                    colspana = cellsa[i].get( '{%s}number-columns-spanned' % self.ns['table'], default=1 )
                    if i < len( cellsb ):
                        colspanb = cellsb[i].get( '{%s}number-columns-spanned' % self.ns['table'], default=1 )
                    else:
                        #this will happen if the repeating row is after a merged column 
                        #causing this row to have fewer cells than the source row. 
                        #The insert index will be the total number of cells, minus 
                        #the (total number of cells spanned minus the cell that will be removed)
                        if len(cellsb) - int( int(span_total)-1 ) < 0:
                            #when there is a merged cell that covers the all the cells.
                            return [ None, 0 ]
                        else:
                            return [ len(cellsb) - int( int(span_total)-1 ), None ]
                    if colspana == colspanb:
                        continue
                    else:
                        span_start = i
                        span_size = colspanb
                        #need to count the total number of columns spanned, so the new cell can be inserted in the correct place
                        span_total += int( span_size )
                    if i == index and span_start is not None:
                        span_end = span_start - 1 + int( span_size )
                        if i >= span_start and i <= span_end:
                            colspan = int(colspanb) + 1
                            fixspanned = span_start
                            return [ None, fixspanned ]
                    elif span_end is not None:
                        if index > span_end:
                            return [ index+1 - int(span_size), None ]
                    elif index < span_start:
                        return [ index, None ]
            return [ index, fixspanned ]
        def _repeatingRow( self, xml, key, params ):
            """
            There is a repeat row modifier on the key, so find the rows, repeat it, then remove 
            the original.
            """
            self.logging.info( "repeating row found." )
            x = self._getXML( xml )
            #perform an xpath search for table cells that contain the repeat row modifier
            rows = x.xpath( '//table:table-row/table:table-cell[contains(.,"%s")]' % key, namespaces=self.ns )
            #rows = x.xpath( '//table:table-row/table:table-cell[contains(.,"~repeatrow")]', namespaces=self.ns )
            if len( rows ):
                table = rows[0].getparent()#.getparent()
                while table.tag != "{%s}table" % self.ns['table']:
                    table = table.getparent()
                if len(table):
                    rowstring = etree.tostring( rows[0].getparent() )
                    #need to escape the key because it contains a pipe (|) which is a special char in regular expressions
                    exp = re.compile( "~%s~" % re.escape( key ) )
                    for v in params:
                        v = self._cleanValue( v )
                        #replace the token with the value
                        replaced = re.sub( exp, "%s" % v, rowstring )
                        row = etree.XML( replaced )
                        #append the new row to the cell
                        table.append( row )
                    #remove the row that contained the original tokens
                    table.remove( rows[0].getparent() )
                self.xml = x
                return etree.tostring( x )
            else:
                return xml
        def _getXML( self, xml ):
            '''
            Getter for the xml property.  This is to help prevent parsing the entire XML 
            document more often than what's necessary.
            '''
            try:
                return etree.XML( xml )
            except:
                return etree.XML( xml.encode( 'utf-8' ) )
        def _getTempFile( self, extension='' ):
            '''
            This is the replacement for os.tmpnam, it should not be suceptable to the
            same symlink injection problems.  I wrapped this simple command in a function 
            beause I don't want to use the file handle it creates. There maybe a more efficent
            way to get a tmp file name, but I havn't found one yet, so I'm using this.
            '''
            file = tempfile.mkstemp( suffix="_pyMMS%s" % extension )
            #I am not going to use this file handle because I found in my research,
            #that I need to open it with the codec module to ensure it's open as utf-8 
            #file intead of whatever the default is
            os.close( file[0] )
            return file[1]
        def _sortparams( self, params ):
            """
            The parameters need to be sorted a little bit; This fixed a problem I had 
            where the repeat section was repeating the section after the conent had 
            already been filled in, and therefore showing the wrong content.
            Sort order is:
                if, repeatsection, repeatcolumn, repeatrow, multiparagraph
            """
            other = []
            modifiers = { 'if':[], 'repeatsection':[],'repeatcolumn':[],'repeatrow':[], 'multiparagraph':[], 'image':[] }
            for key, value in params:
                pipe   = key.find( r"|" )
                colons = key.find( r"::" )  
                if pipe > 0 and colons > 0 and pipe < colons :
                    #sorted.insert( 0, {key:value} )
                    for x in modifiers.keys():
                        if key.find( x+"|" ) > -1:
                            modifiers[x].append( {key:value} )
                else:
                    other.append( {key:value} )
            sorted = []
            sorted.extend( modifiers['if'] )
            sorted.extend( modifiers['repeatsection'] )
            sorted.extend( modifiers['repeatcolumn'] )
            sorted.extend( modifiers['repeatrow'] )
            sorted.extend( modifiers['multiparagraph'] )
            sorted.extend( modifiers['image'] )
            sorted.extend( other )
            return sorted
        def _getFileExtension( self, path ):
            """Get the file extension for the given path"""
            file = splitext(path.lower())
            if len( file ):
                return file[1].replace( '.', '' )
            else:
                return path
        def _cleanValue( self, value ):
            """Sanitize the value to make sure it doesn't break the xml.  The only real problem
            That I have found so far is the & needing to be transformed into &amp; """
            amp = self._getRegEx( 'amp' )
            try:
                return re.sub( amp, u'&amp;', value )
            except:
                return value
class loggingVoid:
    '''This is just a dummy object that does nothing, for when we dont want to log. (unittests,etc.)'''
    def info(self, msg):
        pass
    def error(self, msg):
        pass
    def exception(self, msg):
        pass
    def debug(self, msg):
        pass
#if this module is not being included as a sub module, then start up the soap server
if __name__ == "__main__":
    host = 'localhost'
    port = 8888
    server = SOAPServer( ( host, port ) )
    pyMMS = pyMailMergeService( port=port, host=host )
    namespace = 'urn:approve'
    #I used to just reguister the object and let SOAPpy worry about what methods can and can't get 
    #called, but I found that there was no way to get the SOAPContext if I did that.  So I added 
    #a method to pyMailMerge to return all of the public methods automatically and then they are
    #registered one by one.
    for x in pyMMS.getMethods():
        server.registerFunction( MethodSig( getattr(pyMMS,x), keywords=0, context=1 ), namespace )
    server.serve_forever()