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
@todo: add support for batches of documents, ie a real mail merge
@todo: for the batch support, probably need to be able to support receiving merge values from XML
@todo: unit tests would be good.
@todo: turn into a package, add a setup.py http://www.packtpub.com/article/writing-a-package-in-python
@todo: the openoffice document can have modifiers, like repeatrow| or repeatcolumn| or 
        multiparagraph| It would be really interesting, if these could be plugins or dynamically 
        loaded modules, or something similar
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
from lxml import objectify          #parsing xml
import codecs                       #opening a file for writing as UTF-8
class pyMailMergeService:
        #NOTE: this is not all of the namespaces that are in the document, just the ones I need, plus a few.
        ns = {'table':"urn:oasis:names:tc:opendocument:xmlns:table:1.0", 
              'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0' ,
              'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
              'style':'urn:oasis:names:tc:opendocument:xmlns:style:1.0' }
        xml = None
        def __init__( self ):
            """Create an instance of the document conversion class."""
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
            odtName = param0
            params = param1
            doctype = param2
            '''http://docs.python.org/library/shutil.html'''
            sourcezip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            #create destination zip
            destodt = u"%s.odt" % self._getTempFile()
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
            destpdf = u"%s.%s" % ( self._getTempFile(), doctype ) 
            self.converter.convert( destodt, destpdf )
            f = open( destpdf, 'r' )
            doc = f.read()
            f.close()
            #print destodt
            os.unlink( destodt ) #delete the odt file
            os.unlink( destpdf ) #delete the pdf file
            return base64.b64encode( doc ) #pdf
        def _replaceContent( self, xml, param0 ):
            """
            Do a regular expression find and replace for all of the params, unless the 
            param has multiple values, then do the duplicate rows stuff.
            """
            try:
                params = param0['item'] #no idea why this stuff is inside of 'item'...
            except:
                print "skipping find and replace."
                return xml
            #these were moved out of the loop because compiling the re everytime would be ineffeciant
            para_exp = re.compile( '\<p\>.+\<\/p\>' )
            #for syntax: http://www.php2python.com/wiki/control-structures.foreach/
            for key,value in param0['item']: #param0 is a struct see: http://www.ibm.com/developerworks/webservices/library/ws-pyth16/                
                if type( value ).__name__ == 'instance' or type( value ).__name__ == 'typedArrayType':
                    xml = self._multipleValues( xml, key, value )
                else:
                    if 'multiparagraph|' in key:
                        xml = self._multiparagraph( key, value, xml )
                    #do a simple find and replace with a regular expression
                    exp = re.compile( '~%s~' % re.escape(key) )
                    xml = exp.sub( "%s" % value, xml )
            return xml
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
            for tag in html.findall( 'p' ):
                p = etree.Element( "{%s}p" % self.ns['text'], nsmap=self.ns, attrib=attribs )
                p.text = tag.text
                parent.append( p )
                #add an extra line of blank
                p = etree.Element( "{%s}p" % self.ns['text'], nsmap=self.ns, attrib=attribs )
                parent.append( p )
            parent.remove( paragraphs[0] )
            return etree.tostring( x )
        def _multipleValues( self, xml, key, params ):
            """There are multiple values for this key, so see if there is a modifier on the key to, repeat columns or rows.
            If so, then duplicat the column, or row, then fill in the values.  This is like the normal find and replace, but not 
            global."""
            rowExp = re.compile( "^" + re.escape( 'repeatrow|' ) )
            colExp = re.compile( "^" + re.escape( 'repeatcolumn|' ) )            
            if re.match( rowExp, key ):
                #if it matches the repeat row modifier, then repeat the rows.
                return self._repeatingRow( xml, key, params ) 
            elif re.match( colExp, key ):
                #if it matches the repeat column, token modifer, then repeat the columns.
                return self._repeatingColumn( xml, key, params )
            else:
                self.xml = None #need to clear the xml cache.                
                #if it doesn't match any modifers then, all of the values should already be duplicated, so starting doing find an repalce, but not global
                exp = re.compile( '~%s~' % re.escape(key) )
                for v in params:
                    xml = re.sub( exp, v, xml, count=1 )
                return xml
        def _repeatingColumn( self, xml, key, params ):
            """There is a repeat column modifier on the key, so find the column, repeat it, then remove the original. Also the table has
            an element called table-column, that needs to be updated with the new cell count.  Then any cells that span multiple
            columns needs to be increated by the number of parmas"""
            x = self._getXML( xml )
            cols = x.xpath( '//table:table-row/table:table-cell[contains(.,"%s")]' % key, namespaces=self.ns )
            if len( cols ):
                row = cols[0].getparent()
                if len( row ):
                    table = row.getparent()
                    while table.tag != "{%s}table" % self.ns['table']:
                        table = table.getparent()
                    #for some reason there is a definition in the XML describing how many columns there are, we need to ajust it
                    tablecolumn = table.find( "{%s}table-column" % self.ns['table'] )
                    if tablecolumn is not None:
                        columncount = tablecolumn.get( '{%s}number-columns-repeated' % self.ns['table'] )
                        tablecolumn.set( '{%s}number-columns-repeated' % self.ns['table'], "%s" % ( int( columncount ) + len( params ) - 1 ) )
                    #if there are any columns that have col span, we need to increase that span by the number of params - 1
                    cells = table.findall( '{%s}table-row/{%s}table-cell' % ( self.ns['table'], self.ns['table'] ) );
                    for checkcell in cells:
                        colspan = int( checkcell.get( '{%s}number-columns-spanned' % self.ns['table'], default=0 ) )
                        if colspan > 0:
                            colspan += len( params ) - 1
                            checkcell.set( '{%s}number-columns-spanned' % self.ns['table'], "%s" % colspan )                                        
                    colstring = etree.tostring( cols[0] )
                    #need to escape the key because it contains a pipe (|) which is a special char in regular expressions
                    exp = re.compile( "~%s~" % re.escape( key ) )
                    #columns need to be placed one after the other starting from where the token was found.
                    previouscol = cols[0]
                    for v in params:
                        #replace the token with the value
                        replaced = re.sub( exp, "%s" % v, colstring )
                        col = etree.XML( replaced )
                        #append the new row to the cell
                        previouscol.addnext( col )
                        previouscol = col
                    #remove the row that contained the original tokens
                    row.remove( cols[0] )
                self.xml = etree.tostring( x )
                return self.xml
            else:
                return xml
        def _repeatingRow( self, xml, key, params ):
            """There is a repeat row modifier on the key, so find the rows, repeat it, then remove the original."""
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
            '''Getter for the xml property.  This is to help prevent parsing the entire XML 
            document more often than what's necessary.'''
            #if self.xml is not None:
                #return self.xml
            #else:
            try:
                return etree.XML( xml )
            except:
                return etree.XML( xml.encode( 'utf-8' ) )
        def _getTempFile( self ):
            '''
            This is the replacement for os.tmpnam, it should not be suceptable to the
            same symlink injection problems.  I wrapped this simple command in a function 
            beause I don't want to use the file handle it creates. There maybe a more efficent
            way to get a tmp file name, but I havn't found one yet, so I'm using this.
            '''
            file = tempfile.mkstemp( suffix='_pyMMS' )
            '''I am not going to use this file handle because I found in my research,
            that I need to open it with the codec module to ensure it's open as utf-8 
            file intead of whatever the default is'''
            os.close( file[0] )
            return file[1]
#if this module is not being included as a sub module, then start up the soap server
if __name__ == "__main__":
    server = SOAPServer( ( 'localhost', 8888 ) )
    pyMMS = pyMailMergeService()
    namespace = 'urn:approve'
    server.registerObject( pyMMS, namespace )
    server.serve_forever()