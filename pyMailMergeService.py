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
from sys import path
path.append( '/usr/lib/openoffice.org/program/' )
import base64
import zipfile
import re
import os
from DocumentConverter import *
from SOAPpy import *
from lxml import etree
class soapODConverter:
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
            """Open the odt, and parse for tokens that should be replaced with content, kind 
            of like a mailmerge.  Tokens are single words and start and end with at tilde (~)"""
            '''http://docs.python.org/library/zipfile.html'''
            odtName = param0
            zip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            xml = zip.read( u'content.xml' ) #get the xml string
            #get the tokens from the xml
            '''http://docs.python.org/library/re.html'''
            exp = re.compile( r'~([a-zA-Z\|]+::\w+)~' )
            matches = exp.findall( xml )
            #clean up
            zip.close()
            return matches
        def pdf( self, param0, param1 ):
            """"Convert the odt provided in param0 to a pdf, using the values from param1 
            as mailmerge values."""
            return self.convert( param0, param1, 'pdf' )
        def convert( self, param0, param1, param2='pdf' ):
            """open the odt file which is really just a zip, create a new odt (zip), and copy all
            files from the first odt to the new.  If the file is the content.xml, then 
            replace the tokens from params in the xml before placing it in the new odt.
            Then create a PDF document out of the new odt, and return it base 64 encoded, so 
            it can be transported via Soap.  Soap doesn't seem to like binary data."""
            odtName = param0
            params = param1
            doctype = param2
            '''http://docs.python.org/library/shutil.html'''
            sourcezip = zipfile.ZipFile( u"docs/"+odtName, 'r' )
            #create destination zip
            '''http://docs.python.org/library/os.html#os.tmpfile'''
            #TODO tmpnam is subject to symlink attaks, but since i'm hard coding the tmp dir, it shouldn't be a problem
            destodt = u"%s.odt" % os.tmpnam()
            destzip = zipfile.ZipFile( destodt, 'w' )
            #copy all file from the source zip(odt) the the destination zip
            for x in sourcezip.namelist():
                if x == 'content.xml':
                    tmp = self.replaceContent( sourcezip.read( x ), params )
                else:
                    tmp = sourcezip.read( x )
                destzip.writestr( x, tmp )
            #clan up
            destzip.close()
            sourcezip.close()
            #now convert the odt to pdf
            destpdf = u"%s.%s" % ( os.tmpnam(), doctype ) 
            self.converter.convert( destodt, destpdf )
            f = open( destpdf, 'r' )
            doc = f.read()
            f.close()
            os.unlink( destodt ) #delete the odt file
            os.unlink( destpdf ) #delete the pdf file
            return base64.b64encode( doc )#pdf
        def replaceContent( self, xml, param0 ):
            """Do a regular expression find and replace for all of the params, unless the 
            param has multiple values, then do the duplicate rows stuff."""
            #no idea why this stuff is inside of 'item'...
            #for syntax: http://www.php2python.com/wiki/control-structures.foreach/
            for key,value in param0['item']: #param0 is a struct see: http://www.ibm.com/developerworks/webservices/library/ws-pyth16/
                if type( value ).__name__ == 'instance':
                    xml = self.multipleValues( xml, key, value )
                else:
                    exp = re.compile( '~%s~' % key )
                    xml = re.sub( exp, "%s" % value, xml )
            f = open( '/var/www/html/content.xml', 'w' )
            f.write( xml );
            f.close()
            return xml
        def multipleValues( self, xml, key, params ):
            """There are multiple values for this key, so see if there is a modifier on the key to, repeat columns or rows.
            If so, then duplicat the column, or row, then fill in the values.  This is like the normal find and replace, but not 
            global."""
            rowExp = re.compile( "^" + re.escape( 'repeatrow|' ) )
            colExp = re.compile( "^" + re.escape( 'repeatcolumn|' ) )
            if re.match( rowExp, key ):
                #if it matches the repeat row modifier, then repeat the rows.
                return self.repeatingRow( xml, key, params ) 
            elif re.match( colExp, key ):
                #if it matches the repeat column, token modifer, then repeat the columns.
                return self.repeatingColumn( xml, key, params )
            else:
                #if it doesn't match any modifers then, all of the values should already be duplicated, so starting doing find an repalce, but not global
                print key
                exp = re.compile( '~%s~' % key )
                for v in params:
                    xml = re.sub( exp, "%s" % v, xml, 1 )
                self.xml = None #need to clear the xml cache.                 
                return xml
        def repeatingColumn( self, xml, key, params ):
            """There is a repeat column modifier on the key, so find the column, repeat it, then remove the original. Also the table has
            an element called table-column, that needs to be updated with the new cell count.  Then any cells that span multiple
            columns needs to be increated by the number of parmas"""
            x = self.getXML( xml )
            cols = x.xpath( '//table:table-row/table:table-cell[contains(.,"~repeatcolumn")]', namespaces=self.ns )
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
                self.xml = x
                return etree.tostring( x )
            else:
                return xml
        def repeatingRow( self, xml, key, params ):
            """There is a repeat row modifier on the key, so find the rows, repeat it, then remove the original."""
            x = self.getXML( xml )
            #perform an xpath search for table cells that contain the repeat row modifier
            rows = x.xpath( '//table:table-row/table:table-cell[contains(.,"~repeatrow")]', namespaces=self.ns )
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
        def getXML( self, xml ):
            """Getter for the xml property.  This is to help prevent parsing the entire XML 
            document more often than what's necessary."""
            if self.xml is not None:
                return self.xml
            else:
                return etree.XML( xml )
if __name__ == "__main__":
    server = SOAPServer( ( 'localhost', 8888 ) )
    sodc = soapODConverter()
    namespace = 'urn:approve'
    server.registerObject( sodc, namespace )
    server.serve_forever()