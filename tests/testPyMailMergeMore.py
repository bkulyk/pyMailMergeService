#!/usr/bin/env python
from sys import path
path.append( '..' )
from mms import pyMailMerge
from mms.OfficeDocument import *
from mms.OfficeDocument.WriterDocument import WriterDocument
import unittest
import os

#define unit tests
class testPyMailMergeMore( unittest.TestCase ):
    
    def test_toNumber(self):
        self.assertEqual( 4, WriterDocument.toNumber( '4' ) )
        self.assertEqual( 4000, WriterDocument.toNumber( '4,000.00' ) )
        self.assertEqual( 4000, WriterDocument.toNumber( '$4,000.00' ) )
        self.assertEqual( 4000, WriterDocument.toNumber( '4 000.00 $' ) )
    
    def test_invertTuple(self):
        
        data = [ [1, 2, 3, 4, 5],
                 [6, 7, 8, 9, 10] ]
        
        expected = ( ( 1, 6 ), 
                     ( 2, 7 ), 
                     ( 3, 8 ), 
                     ( 4, 9 ), 
                     ( 5, 10 ) )
        
        self.assertEquals( expected, WriterDocument.invertTuple( data   ) )
    
    def test_updateChart(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/charttest.odt' ) )
        outFile = os.path.join( os.path.dirname( __file__ ), 'docs/charttest.out.pdf' )
        pmm = pyMailMerge( path ) 
        
        xml = """<tokens>
            <token>
                <name>repeatrow|timeperiod</name>
                <value>Year - 2009</value>
                <value>Year - 2010</value>
                <value>Year - 2011</value>
                <value>Year - 2012</value>
            </token>
            <token>
                <name>premium</name>
                <value>$505</value>
                <value>$1500</value>
                <value>$1000</value>
                <value>$1000</value>
            </token>
            <token>
                <name>paidclaims</name>
                <value>$200</value>
                <value>$2000</value>
                <value>$50</value>
                <value>$3000</value>
            </token>
            <token>
                <name>lossratio</name>
                <value>34%</value>
                <value>33%</value>
                <value>35%</value>
                <value>35%</value>
            </token>
        </tokens>"""
        
        x = pyMailMerge._readParamsFromXML( xml )
        pmm._process( x )
        pmm.document.refresh()
        pmm.document.saveAs( outFile )
        
        self.assertTrue( True )

if __name__ == '__main__':
    unittest.main()