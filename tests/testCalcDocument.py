#!/usr/bin/env python
from sys import path, exit
path.append( '..' )
path.append( '../src' )
path.append( '../src/OfficeDocument' )
from OfficeDocument.CalcDocument import *
import unittest, os, uno

#define unit tests
class testWriterDocument( unittest.TestCase ):
    def setUp(self):
        pass

#    def test_duplicateRow(self):
#        ood = CalcDocument()
#        ood.open( '/usr/share/pyMailMergeService/tests/docs/repeat_row_repeat_column.odt' )
#        ood.duplicateRow( "~a2~", True )
#        ood.saveAs( '/usr/share/pyMailMergeService/tests/docs/repeat_row_repeat_column.pdf' )
#        ood.close()

    def test_getNamedRanges( self ):
        doc = CalcDocument()
        path = os.path.dirname( os.path.abspath( __file__ ) ) + "/docs/getNamedRanges.ods"
        doc.open( path )
        names = doc.getNamedRanges()
        self.assertTrue( 'first' in names )
        self.assertTrue( 'myranges' in names )
        self.assertTrue( 'cola' in names )
        self.assertTrue( 'colb' in names )
        self.assertTrue( 'colc' in names )
        doc.close()
    
    def test_getNamedRangeData( self ):
        doc = CalcDocument()
        path = os.path.dirname( os.path.abspath( __file__ ) ) + "/docs/getNamedRanges.ods"
        doc.open( path )
        #test a couple of named ranges
        values = doc.getNamedRangeData( 'first' )
        self.assertEquals( (1, 2, 3), values[0] )
        self.assertEquals( (4, 40718, 6), values[1] )
        self.assertEquals( (7, 8, 9), values[2] )
        #another named range
        values = doc.getNamedRangeData( 'cola' )
        self.assertEquals( (1, 4, 7, 10), values )
        doc.close()
                
    def test_getRangeStrings( self ):
        doc = CalcDocument()
        path = os.path.dirname( os.path.abspath( __file__ ) ) + "/docs/getNamedRanges.ods"
        doc.open( path )
        values = doc.getNamedRangeStrings( 'first' )
        doc.close()
        expected = [ [ '1', '2', '3' ],
                     [ '4', 'Friday, June 24, 2011', '6' ],
                     [ '7', '8', '9' ],
                     [ '10', '11', '12' ] ]
        self.assertEqual( expected, values )

if __name__ == "__main__":
    unittest.main()