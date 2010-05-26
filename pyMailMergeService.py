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
@complete: need to be able to have a modifier for an IF statement, one of our documents needs to 
       be able to omit a section depending on a varaible that would come through the webservice
       client.
@complete: Modifier for repeating section.
@todo: add support for batches of documents, ie a real mail merge
@todo: for the batch support, probably need to be able to support receiving merge values from XML or JSON maybe
@todo: unit tests would be good.
@todo: turn into a package, add a setup.py http://www.packtpub.com/article/writing-a-package-in-python
@todo: the openoffice document can have modifiers, like repeatrow| or repeatcolumn| or 
       multiparagraph| It would be really interesting, if these could be plugins or dynamically 
       loaded modules, or something similar
@todo: Before writing final xml files to the new odt. Make sure there are not any macros. Marcos could cause harm to the system.
@todo: Add some logging information.  
        What types of documents are being converted
        What IPs are using the service, simple stats about client computer
"""
from sys import path
path.append( '/usr/lib/openoffice.org/program/' )
import base64                       #to be able to return the output file utf encoded
import zipfile                      #to open odt so we can do find and replace on the xml
import re                           #regular expressions for doing find and replace on the xml
import os                           #removing the temp files
import tempfile                     #create temp files for writing xml and output
from DocumentConverter import *     #to convert open office documents
from SOAPpy import *                #soap interface
from lxml import etree              #parsing xml
import codecs                       #opening a file for writing as UTF-8
class pyMailMergeService:
        #NOTE: this is not all of the namespaces that are in the document, just the ones I need, plus a few.
        ns = {'table':"urn:oasis:names:tc:opendocument:xmlns:table:1.0", 
              'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0' ,
              'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
              'style':'urn:oasis:names:tc:opendocument:xmlns:style:1.0' }
        xml = None
        _params = None
        def __init__( self, connect=True ):
            """Create an instance of the document conversion class."""
            if connect:
                self.converter = DocumentConverter()
        def hello( self, param0 ):
            """This is a simple 'Hello World' function to test soap calls"""
            return u"hello %s " % param0
        def getTokens( self, param0 ):
            """
            Open the odt, and parse for tokens that should be replaced with content, kind 
            of like a mailmerge.  Tokens are single words and start and end with at tilde (~)
            """
            '''http://docs.python.org/library/zipfile.html'''
            odtName = param0
            zip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            #get the tokens from the xml
            '''http://docs.python.org/library/re.html'''
            exp = re.compile( r'~([a-zA-Z\|]+::\w+\|?\w*)~' )
            matches = []
            #I found that I need to look for tokens in the styles and meta fils as well, because meta has the document title, and styles has the content for the document footers and probably headers
            for file in [ u'content.xml', u'styles.xml', u'meta.xml' ]:
                xml = zip.read( file )
                matches.extend( exp.findall( xml ) )
            #clean up
            zip.close()
            matches = list( set( matches ) ) #removes duplicate items from the list
            return matches
        def pdf( self, param0, param1 ):
            """"Convert the odt provided in param0 to a pdf, using the values from param1 
            as mailmerge values."""
            return self.convert( param0, param1, 'pdf' )
        def convert( self, param0, param1, param2='pdf' ):
            """
            open the odt file which is really just a zip, create a new odt (zip), and copy all
            files from the first odt to the new.  If the file is the content.xml, then 
            replace the tokens from params in the xml before placing it in the new odt.
            Then create a PDF document out of the new odt, and return it base 64 encoded, so 
            it can be transported via Soap.  Soap doesn't seem to like binary data.
            """
            """NEED to clear the params because they are cached from xml file to xml file 
            because the tokens would be the same for, content.xml as styles.xml, etc.  But 
            it was remembering these from document to document (very bad)"""
            self._params = None
            odtName = param0
            params  = param1
            doctype = param2
            print "Converting file: %s to type: %s" % ( odtName, doctype )
            '''http://docs.python.org/library/shutil.html'''
            sourcezip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            #create destination zip
            destodt = self._getTempFile( '.odt' )
            destzip = zipfile.ZipFile( destodt, 'w' )
            #copy all file from the source zip(odt) the the destination zip
            for x in sourcezip.namelist():
                if x in ['content.xml', 'meta.xml', 'styles.xml']:
                    tmp = self._replaceContent( sourcezip.read( x ), params )
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
                else:
                    tmp = sourcezip.read( x )
                    destzip.writestr( x, tmp )
            #clan up
            destzip.close()
            sourcezip.close()
            #now convert the odt to pdf
            destpdf = u"%s" % ( self._getTempFile( "."+doctype ) )
            converter = DocumentConverter()
            converter.convert( destodt, destpdf )
            converter = None
            f = open( destpdf, 'r' )
            doc = f.read()
            f.close()
            os.unlink( destodt ) #delete the odt file
            os.unlink( destpdf ) #delete the pdf file
            return base64.b64encode( doc ) #pdf
        def _replaceContent( self, xml, param0 ):
            """
            Do a regular expression find and replace for all of the params, unless the 
            param has multiple values, then do the duplicate rows stuff.
            """
            if self._params is None: #preventing having to soft the params more than once.
                try:
                    params = param0['item'] #no idea why this stuff is inside of 'item'...
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
                        print "skipping find and replace."
                        return xml
                self._params = params
            else:
                params = self._params
            #these were moved out of the loop because compiling the re everytime would be ineffeciant
            para_exp = re.compile( '\<p\>.+\<\/p\>' )
            #for syntax: http://www.php2python.com/wiki/control-structures.foreach/
            for dictionary in params:
                key = dictionary.keys()[0]
                value = dictionary.values()[0]
                if type( value ).__name__ in ('instance', 'list', 'typedArrayType' ):
                    xml = self._multipleValues( xml, key, value )
                else:
                    if r'multiparagraph|' in key:
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
                    xml = exp.sub( "%s" % value, xml )
            return xml
        def _if( self, key, value, xml ):
            """
            If there is a modifier tag for an if statement, only show the stuff between the if
            and the endif modifiers IF the value is 1, '1' or true.
            This piece is really touchy as the start and end tags have to be in the same 
            depth of xml nodes, which can be difficult.  This whole piece may need be be
            reworked to do the editing via XML instead of text
            """
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
                if value == 0:
                    print "\"if\" modifier failed, reverting to xml before statement"
                return xmlbackup
            """The removal of content in the manner that it was just done could break the 
            xml structure and cause it to not parse anymore.  That would be bad.  So if the 
            xml is broken then don't use this modified version and fall back to the old."""
            try:
                xml = etree.XML( tmp )
                return tmp
            except:
                print "xml is bad reverting"
                return xmlbackup
        def _repeatsection( self, key, value, xml ):
            """
            Repeat a section of the document as many times as the value says, which should be 
            and integer.
            """
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
            """The removal of content in the manner that it was just done could break the 
            xml structure and cause it to not parse anymore.  That would be bad.  So if the 
            xml is broken then don't use this modified version and fall back to the old."""
            try:
                tmp = etree.XML( xml )
                return xml
            except:
                return xmlbackup
        def _multiparagraph( self, key, value, xml ):
            """
            If there was a single paragraph containing the merge tag, and now there 
            needs to be multiple paragrpahs, convert html <p> tags to their openoffice
            couterparts, which is still a p, but with a namespace and style attribute
            """
            self.xml = None
            x = self._getXML( xml )
            paragraphs = x.xpath( '//text:p[contains(text(),"~%s~")]' % key, namespaces=self.ns )
            if not len( paragraphs ):
                return xml
            parent = paragraphs[0].getparent()
            #need to add the paragraphy style attribute (and any others) to the new p tags
            attribs = paragraphs[0].attrib
            html = etree.XML( "<html>" + value + "</html>" )
            previous = paragraphs[0]
            for tag in html.findall( 'p' ):
                p = etree.Element( "{%s}p" % self.ns['text'], nsmap=self.ns, attrib=attribs )
                p.text = tag.text
                #need to add next instead of append because it was messing up ordering of content
                previous.addnext( p )
                previous = p
                #add an extra line of blank
                p = etree.Element( "{%s}p" % self.ns['text'], nsmap=self.ns, attrib=attribs )
                previous.addnext( p )
                previous = p
            parent.remove( paragraphs[0] )
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
                #if it matches the repeat column, token modifer, then repeat the columns.
                return self._repeatcolumn( xml, key, params )
            else:
                self.xml = None #need to clear the xml cache.                
                #if it doesn't match any modifers then, all of the values should already be duplicated, so starting doing find an repalce, but not global
                exp = re.compile( '~%s~' % re.escape(key) )
                for v in params:
                    xml = re.sub( exp, v, xml, count=1 )
                return xml
        def _repeatcolumn( self, xml, key, params ):
            """
            Repeat an entire column of a given table, replacing values as we go.
            """
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
                        '''need to hold off on repeating the cells in the source column
                        until last because it would mess up the column count that I need
                        in order to make the merged cell count stuff work''' 
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
                newElement = oldString.replace( "~%s~" % key, param )
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
                        tablecolumn = tablecolumns[0] #should only be one.
                        tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                    else: #as a last option, duplicate the first table-column
                        tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                        tablecolumn = tablecolumns[0]
                    tablecolumnString = etree.tostring( tablecolumn )
                    tablecolumns = row.getparent().xpath( r"./table:table-column", namespaces=self.ns )
                    if len( cells ) < int( len( cells ) + i ):
                        tablecolumns[ len( tablecolumns ) -1 ].addnext( etree.XML( tablecolumnString ) )
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
                        '''this will happen if the repeating row is after a merged column 
                        causing this row to have fewer cells than the source row. 
                        The insert index will be the total number of cells, minus 
                        the (total number of cells spanned minus the cell that will be removed)'''
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
            '''I am not going to use this file handle because I found in my research,
            that I need to open it with the codec module to ensure it's open as utf-8 
            file intead of whatever the default is'''
            os.close( file[0] )
            return file[1]
        def _sortparams( self, params ):
            """
            The parameters need to be sorted a little bit; This fixed a problem I had 
            where the repeat section was repeating the section after the conent had 
            already been filled in, and therefore showing the wrong content.
            Sort order is:
                if, repeatsection, repeatcolumn, repeatrow
            """
            other = []
            modifiers = { 'if':[], 'repeatsection':[],'repeatcolumn':[],'repeatrow':[]}
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
            for v in modifiers.values():
                sorted.extend( v )
            sorted.extend( other )
            return sorted
#if this module is not being included as a sub module, then start up the soap server
if __name__ == "__main__":
    server = SOAPServer( ( 'localhost', 8888 ) )
    pyMMS = pyMailMergeService()
    namespace = 'urn:approve'
    server.registerObject( pyMMS, namespace )
    server.serve_forever()