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
    def test_updateChart(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/charttest.odt' ) )
        outFile = os.path.join( os.path.dirname( __file__ ), 'docs/charttest.out.pdf' )
        pmm = pyMailMerge( path ) 
        
        xml = """<tokens>
            <token>
                <name>repeatrow|timeperiod</name>
                <value>2009</value>
                <value>2010</value>
                <value>2011</value>
                <value>2012</value>
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
                <value>$50</value>
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
#        pmm.document.updateCharts()
        pmm.document.saveAs( outFile )
        
        self.assertTrue( True )
        
        
#        results = self._getFirstTableData( outFile )
#        
#        expected = [[ 'First', "Third" ],
#                    [ 'A', 'C' ],
#                    [ 'D', 'F' ] ]
#        
#        self.assertEqual( expected, results )

if __name__ == '__main__':
    unittest.main()