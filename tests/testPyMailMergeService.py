#!/usr/bin/env python
from sys import path
path.append( '..' )
from pyMailMergeService import *
import unittest
import zipfile
from lxml import etree
#define unit tests
class testPyMailMergerService( unittest.TestCase ):
    def setUp(self):
        pass
    def test_getXML( self ):
        xml = getXML( "simple_repeat_column.odt", True )
        self.assertEqual( '_Element', type( xml ).__name__ )
        xml = getXML( "simple_repeat_column.odt" )
        self.assertEqual( 'str', type( xml ).__name__ )
    def test_getTableText(self):
        #mock data
        xml = getXML( "simple_repeat_column.odt" )
        #assert that the ODT hasn't changed and that the xml parsing is working
        matrix = getTableText( xml )
        self.assertEquals( '~repeatcolumn|token::test~', matrix[0][0] )
        self.assertEquals( 2, len( matrix ) )
        self.assertEquals( 2, len( matrix[0] ) )
        self.assertEquals( 2, len( matrix[1] ) )
        self.assertEquals( "a2", matrix[0][1] )
        self.assertEquals( "b1", matrix[1][0] )
        self.assertEquals( "b2", matrix[1][1] )
    def test_simple_repeat_column( self ):
        #mock data
        xml = getXML( "simple_repeat_column.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( False )        
        xml = pymms._repeatcolumn( xml, key, value )
        #test results
        matrix = getTableText( xml )
        self.assertEquals( 2, len( matrix ) )
        self.assertEquals( 3, len( matrix[0] ) )
        self.assertEquals( 3, len( matrix[1] ) )
        self.assertEquals( "replace1", matrix[0][0] )
        self.assertEquals( "replace2", matrix[0][1] )
        self.assertEquals( "a2", matrix[0][2] )
        self.assertEquals( "b1", matrix[1][0] )
        self.assertEquals( "b1", matrix[1][1] )
        self.assertEquals( "b2", matrix[1][2] )
    def test_repeat_column_merged_columns_before( self ):
        xml = getXML( "repeat_column_merged_columns_above.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( False )
        xml = pymms._repeatcolumn( xml, key, value )
        matrix = getTableText( xml )
        self.assertEqual( 'a1', matrix[0][0] )
        self.assertEqual( 'a2', matrix[0][1] )
        self.assertEqual( 'a2', matrix[0][2] )
        self.assertEqual( 'a3', matrix[0][3] )
        self.assertEqual( 'a4', matrix[0][4] )
        self.assertEqual( "b1", matrix[1][0] )
        self.assertEqual( "b2b3", matrix[1][1] )
        self.assertEqual( 'b4', matrix[1][2] )
        self.assertEqual( 'c1', matrix[2][0] )
        self.assertEqual( 'replace1', matrix[2][1] )
        self.assertEqual( 'replace2', matrix[2][2] )
        self.assertEqual( 'c3', matrix[2][3] )
        self.assertEqual( 'c4', matrix[2][4] )
        xml = etree.XML( xml )
    def test_repeat_column_merged_columns_after( self ):
        xml = getXML( "repeat_column_merged_columns_before.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( False )
        xml = pymms._repeatcolumn( xml, key, value )
        matrix = getTableText( xml )
        self.assertEqual( 'a1', matrix[0][0] )
        self.assertEqual( 'a2', matrix[0][1] )
        self.assertEqual( 'replace1', matrix[0][2] )
        self.assertEqual( 'replace2', matrix[0][3] )
        self.assertEqual( 'b1b2', matrix[1][0] )
        self.assertEqual( 'b3', matrix[1][1] )
        self.assertEqual( 'b3', matrix[1][2] )
        self.assertEqual( 'c1', matrix[2][0] )
        self.assertEqual( 'c2', matrix[2][1] )
        self.assertEqual( 'c3', matrix[2][2] )
        self.assertEqual( 'c3', matrix[2][3] )
def getTableText( xml ):
    ns = {'table':"urn:oasis:names:tc:opendocument:xmlns:table:1.0", 
              'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0' ,
              'office':'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
              'style':'urn:oasis:names:tc:opendocument:xmlns:style:1.0' }
    xml = etree.XML( xml )
    matrix = []
    rows = xml.xpath( "//table:table-row", namespaces=ns )
    for row in rows:
        items = []
        cells = row.xpath( "./table:table-cell/text:p", namespaces=ns )
        for p in cells:
            items.append( p.text )
        matrix.append( items )
    return matrix
def getXML( filename, parsed=False ):
    zip = zipfile.ZipFile( u"docs/"+filename, 'r' )
    if parsed:
        xml = etree.XML( zip.read( "content.xml" ) )
    else:
        return zip.read( "content.xml" )
    return xml
#testtable
if __name__ == '__main__':
    unittest.main()
