#!/usr/bin/env python
from sys import path
path.append( '..' )
from pyMailMerge import *
import unittest
import os
#define unit tests
class testPyMailMerge( unittest.TestCase ):
    xml = r'''<tokens>
            <token>
                <name>fake::token</name>
                <value>value1</value>
            </token>
            <token>
                <name>fake::array</name>
                <value>first</value>Inspect
                <value>second</value>
            </token>
            <token>
                <name>token::fake</name>
                <value>value2</value>
            </token>
        </tokens>'''
    def setUp(self):
        pass
    def test_sortParams(self):
        x = pyMailMerge._sortParams( [
                                   { 'token':'fake::token','value':'value' }, 
                                   { 'token':'html|fake::withhtml','value':'<h1>whatever</h1>' }, 
                                   { 'token':'repeatrow|fake::repeatingrow','value':['row1','row2'] }, 
                                   { 'token':'fake::tokens','value':['1','2'] }, 
                                   { 'token':'if|fake::ifstatement', 'value':'1' } 
        ] )
        self.assertEquals( 'if|fake::ifstatement', x[0]['token'] )
        self.assertEquals( 'repeatrow|fake::repeatingrow', x[1]['token'] )
        self.assertEquals( 'html|fake::withhtml', x[2]['token'] )
        #don't care about the other two tokens as long as the order of the ones with modifiers is correct
        self.assertEquals( 5, len( x ) )
    def test_readParamsFromXML(self):
        x = pyMailMerge._readParamsFromXML( self.xml )
        self.assertEquals( 'fake::token', x[0]['token'] )
        self.assertEquals( 'value1', x[0]['value'] )
        self.assertEquals( 'fake::array', x[1]['token'] )
        self.assertEquals( 'first', x[1]['value'][0] )
        self.assertEquals( 'second', x[1]['value'][1] )
        self.assertEquals( 'token::fake', x[2]['token'] )
        self.assertEquals( 'value2', x[2]['value'] )
    def test_process(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.odt' ) )
        pmm = pyMailMerge( path )
        x = [
               { 'token':'company::name','value':'Random Company' }, 
               { 'token':'company::address1','value':'123 Mystery Lane' }, 
               { 'token':'company::city','value':"Winnipeg" }, 
               { 'token':'company::prov','value':"MB" }, 
               { 'token':'company::phone', 'value':'123-555-4567' },
               { 'token':"client::company", "value":"Client Company" },
               { 'token':'client::name', 'value':'Random Dude' },
               { 'token':'client::city', 'value':'Brandon' },
               { 'token':'client::prov', 'value':'MB' },
               { 'token':'client::postalcode', 'value':'R3J 2U8' },
               { 'token':'repeatrow|product::desc', 'value':['SKU 123 - Hammer','SKU 223 - Nail'] },
               { 'token':'product::rate', 'value':['6.98','1.99'] },
               { 'token':'product::qty', 'value':[ '1', '10' ] },
               { 'token':'product::total', 'value':[ '6.98', '19.90' ] },
               { 'token':'if|paid', 'value':'1' },
               { 'token':'if|notpaid', 'value':'0' },
               { 'token':'paid::date', 'value':'Jan 01, 3011' },
               { 'token':'payment::due', 'value':'Feb 01, 3011' },
               { 'token':'paid', 'value':'PAID' }
        ]
        pmm._process( x )
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.out.odt' ) )
if __name__ == '__main__':
    unittest.main()