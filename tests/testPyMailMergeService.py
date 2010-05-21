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
    def test_simple_repeat_column( self ):
        xml = getXML( "simple_repeat_column.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replaceA", "replaceB" )
        pymms = pyMailMergeService() 
        pymms._repeatcolumn( xml, key, value )
        
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
