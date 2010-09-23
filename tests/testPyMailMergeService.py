#!/usr/bin/env python
from sys import path
path.append( '..' )
from pyMailMergeService import *
import unittest
import zipfile
import os
from lxml import etree
#define unit tests
class testPyMailMergerService( unittest.TestCase ):
    def setUp(self):
        pass
    def test_hello(self):
        pymms = pyMailMergeService( enablelogging=False )
        self.assertEquals( 'hello test', pymms.hello( 'test' ) )
        self.assertEquals( 'hello world', pymms.hello() )
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
        pymms = pyMailMergeService( enablelogging=False )
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
    def test_repeat_column_merged_columns_above( self ):
        xml = getXML( "repeat_column_merged_columns_above.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( enablelogging=False )
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
    def test_repeat_column_merged_columns_before( self ):
        xml = getXML( "repeat_column_merged_columns_before.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( enablelogging=False )
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
    def test_repeat_column_merged_columns_before_2( self ):
        xml = getXML( "repeat_column_merged_columns_before_2.odt" )
        key = "repeatcolumn|token::test"
        value = ( "replace1", "replace2" )
        #run methods
        pymms = pyMailMergeService( enablelogging=False )
        xml = pymms._repeatcolumn( xml, key, value )
        matrix = getTableText( xml )
        self.assertEqual( 'a1', matrix[0][0] )
        self.assertEqual( 'a2', matrix[0][1] )
        self.assertEqual( 'a3', matrix[0][2] )
        self.assertEqual( 'replace1', matrix[0][3] )
        self.assertEqual( 'replace2', matrix[0][4] )
        self.assertEqual( 'b1b2', matrix[1][0] )
        self.assertEqual( 'b3', matrix[1][1] )
        self.assertEqual( 'b4', matrix[1][2] )
        self.assertEqual( 'b4', matrix[1][3] )
        self.assertEqual( 'c1', matrix[2][0] )
        self.assertEqual( 'c2', matrix[2][1] )
        self.assertEqual( 'c3', matrix[2][2] )
        self.assertEqual( 'c4', matrix[2][3] )
        self.assertEqual( 'c4', matrix[2][4] )
    def test_get_file_extension(self):
        pymms = pyMailMergeService( enablelogging=False )
        self.assertEqual( 'odt', pymms._getFileExtension( "simple_repeat_column.odt" ) )
        self.assertEqual( 'odt', pymms._getFileExtension( "invoice.odt" ) )
    def test_if_section_simple_true(self):
        xml = getXML( "if_section_simple.odt" )
        key = "if|token::testif"
        value = 1
        #run method
        pymms = pyMailMergeService( enablelogging=False )
        xml = pymms._if( key, value, xml )
        xml = etree.XML( xml )
        #test values
        p = xml.xpath( "//text:p[contains(.,'%s')]" % "blah blah", namespaces={'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0'} )
        self.assertEquals( 1, len( p ) )
        self.assertTrue( p[0].text.find( 'blah blah' ) > -1 )
    def test_if_section_simple_false(self):
        xml = getXML( "if_section_simple.odt" )
        key = "if|token::testif"
        value = 0
        #run method
        pymms = pyMailMergeService( enablelogging=False )
        xml = pymms._if( key, value, xml )
        xml = etree.XML( xml )
        #test values
        p = xml.xpath( "//text:p[contains(.,'%s')]" % "blah blah", namespaces={'text':'urn:oasis:names:tc:opendocument:xmlns:text:1.0'} )
        self.assertEqual( 0, len( p ) )
    def test_sortparams(self):
        params = [ ['test::token', 0], ['if|test::if', 1], ['repeatcolumn|test::repeatcol', 2], ['repeatrow|test::repeat_row', 3], ['repeatsection|test::repeatsect', 4] ]
        #params[ "test::token" ] = ['0']
        #run method
        pymms = pyMailMergeService( enablelogging=False )
        params = pymms._sortparams( params )
        #test the output
        self.assertEquals( 1, params[0]['if|test::if'] )
        self.assertEquals( 4, params[1]['repeatsection|test::repeatsect'] )
        self.assertEquals( 2, params[2]['repeatcolumn|test::repeatcol'] )
        self.assertEquals( 3, params[3]['repeatrow|test::repeat_row'] )
        self.assertEquals( 0, params[4]['test::token'] )
        self.assertEquals( 5, len( params ) )
    def test_AmpRegEx(self):
        pymms = pyMailMergeService( enablelogging=False )
        amp = pymms._getRegEx( 'amp'  )
        self.assertEquals( "blah &amp; blah", re.sub( amp, "&amp;", "blah & blah" ) )
        self.assertEquals( "blah &amp; blah &amp;", re.sub( amp, "&amp;", "blah & blah &" ) )
        self.assertEquals( "blah &amp; blah", re.sub( amp, "&amp;", "blah &amp; blah" ) )
        self.assertEquals( "blah &lt; blah", re.sub( amp, "&amp;", "blah &lt; blah" ) )
        self.assertEquals( "blah &#8226; blah", re.sub( amp, "&amp;", "blah &#8226; blah" ) )
    def test_TokenRegEX(self):
        pymms = pyMailMergeService( enablelogging=False )
        token = pymms._getRegEx( 'tokens' )
        sampleStringOfPossibleTokens = """
        ~blah::blah~     #should be found
        ~blah::blah|1~   #should be found
        ~object::method|param~    #should be found
        ~modifier|object::method|param~    #should be found
        ~modifier|object::method|param|param2~    #should be found
        ~object::method2~    #should be found
        ~object::method|param|paramwith)~    #shold be found -- this is a specific use case for my company... I know it's a little weird.
        ~object::rightnext1~~object::rightnext2~ #should both be found
        ~object::noclosingtilde    #should NOT work because it has no closing tilde
        object::noopeningtilde~    #should NOT work because it has no opening tilde
        ~::blah|12~      #should NOT be found because the object part is missing
        ~object::method_test~    #should NOT be found
        ~modifier|object::method|para#~    #should NOT be found
        blahblah~object::mixedinwithothertext~blahblah  I'm just proving that this should be there because it'sm mixed in with other text
        ~object1::method1~    #should NOT be found because it has a number in the object, which really isn't a problem unless it's first but whatever.
        ~~    #should NOT be found
        """
        matches = token.findall( sampleStringOfPossibleTokens )
        self.assertTrue( "::blah|12" not in matches )
        self.assertTrue( "blah::blah" in matches )
        self.assertTrue( "blah::blah|1" in matches )
        self.assertTrue( "object::method|param" in matches )
        self.assertTrue( "modifier|object::method|param" in matches )
        self.assertTrue( "modifier|object::method|param|param2" in matches )
        self.assertTrue( "modifier|object::method|para#" not in matches )
        self.assertTrue( "object::method2" in matches )
        self.assertTrue( 'object::mixedinwithothertext' in matches )
        self.assertTrue( 'object::rightnext1' in matches )
        self.assertTrue( 'object::rightnext2' in matches )
        self.assertTrue( 'object::modifier_test' not in matches )
        self.assertTrue( 'object::method|param|paramwith)' in matches )
        self.assertTrue( 'object::noclosingtilde' not in matches )
        self.assertTrue( 'object::noopeningtilde' not in matches )
        self.assertTrue( 'object1::method1' not in matches )
        self.assertTrue( "" not in matches )
    def testLogging(self):
        try:
            #remove the file in case it exists before the test.
            os.unlink( 'pymms.log' )
        except:
            pass
        pmms = pyMailMergeService( enablelogging=False )
        pmms.hello( 'world' )
        #I just want to make sure the log file does not get created, if the log 
        #does not exist then an exception will be thrown.
        file = None
        try:
            file = open( 'pymms.log' )
        except:
            pass
        self.assertEquals( None, file )
    def testOSPathExists(self):
        self.assertTrue( os.path.exists( 'docs/if_section_simple.odt' ) )
        self.assertFalse( not os.path.exists( 'docs/if_section_simple.odt' ) )
    def testParseXMLParams(self):
        xml = """<?xml version="1.0" encoding="UTF-8"?>
        <tokens>
            <token>
                <name>whatever::whatever</name>
                <value>somevalue</value>
            </token>
            <token>
                <name>token::multivalue</name>
                <value>a</value>
                <value>b</value>
                <value>c</value>
            </token>
        </tokens>"""
        pymms = pyMailMergeService( enablelogging=False )
        params = pymms._parseXMLParams( xml )
        params = pymms._sortparams( params )
        self.assertEquals( "somevalue", params[0]['whatever::whatever'] )
        self.assertEquals( "a", params[1]['token::multivalue'][0] )
        self.assertEquals( "b", params[1]['token::multivalue'][1] )
        self.assertEquals( "c", params[1]['token::multivalue'][2] )
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
if __name__ == '__main__':
    unittest.main()
