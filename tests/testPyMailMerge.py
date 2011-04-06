#!/usr/bin/env python
from sys import path
path.append( '../src' )
path.append( '../src/lib' )
from pyMailMerge import *
import unittest
import os
#define unit tests
class testPyMailMerge( unittest.TestCase ):
    """
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
               { 'token':'product::total', 'value':['6.98', '19.90' ] },
               { 'token':'if|paid', 'value':'1' },
               { 'token':'if|notpaid', 'value':'0' },
               { 'token':'paid::date', 'value':'Jan 01, 3011' },
               { 'token':'payment::due', 'value':'Feb 01, 3011' },
               { 'token':'paid', 'value':'PAID' },
               { 'token':'html|notes', 'value':'<div style="font-family: Arial;"><h3>Terms:</h3><ol><li>Payment due in <strong>30</strong> days.</li><li>No refunds</li></ol></div>' },
               { 'token':'repeatsection|repeater', 'value':'2' }
        ]
        pmm._process( x )
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.out.odt' ) )
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.out.pdf' ) )
    def test_getTokens(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.odt' ) )
        pmm = pyMailMerge( path )
        x = pmm.getTokens()
        x.sort()
        self.assertEquals( 0,  x.index( '~client::city~' ) ) 
        self.assertEquals( 1,  x.index( '~client::company~' ) ) 
        self.assertEquals( 2,  x.index( '~client::name~' ) )
        self.assertEquals( 3,  x.index( '~client::postalcode~' ) ) 
        self.assertEquals( 4,  x.index( '~client::prov~' ) )
        self.assertEquals( 5,  x.index( '~company::city~' ) )
        self.assertEquals( 6,  x.index( '~company::name~' ) )
        self.assertEquals( 7,  x.index( '~company::phone~' ) )
        self.assertEquals( 8,  x.index( '~company::prov~' ) )
        self.assertEquals( 9,  x.index( '~endif|notpaid~' ) )
        self.assertEquals( 10, x.index( '~endif|paid~' ) )
        self.assertEquals( 11, x.index( '~html|notes~' ) )
        self.assertEquals( 12, x.index( '~if|notpaid~' ) )
        self.assertEquals( 13, x.index( '~if|paid~' ) )
        self.assertEquals( 14, x.index( '~paid::date~' ) ) 
        self.assertEquals( 15, x.index( '~payment::due~' ) ) 
        self.assertEquals( 16, x.index( '~product::qty~' ) )
        self.assertEquals( 17, x.index( '~product::rate~' ) )
        self.assertEquals( 18, x.index( '~product::total~' ) )
        self.assertEquals( 19, x.index( '~repeatrow|product::desc~' ) ) 
    def testRepeatSection(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSection.odt' ) )
        pmm = pyMailMerge( path )
        x = [
               { 'token':'repeatsection|first', 'value':'4' }
        ]
        pmm._process( x )
        pmm.document.refresh()
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSection.out.pdf' ) )
    def testRepeatSectionReadOnly(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSection_readonly.odt' ) )
        pmm = pyMailMerge( path )
        x = [
               { 'token':'repeatsection|first', 'value':'4' }
        ]
        pmm._process( x )
        pmm.document.refresh()
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSection_readonly.out.pdf' ) )
        """
    def testRepeatSectionTable(self):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSectionTable.odt' ) )
        pmm = pyMailMerge( path )
        x = [
               { 'token':'repeatsection|first', 'value':'4' }
        ]
        pmm._process( x )
        pmm.document.refresh()
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/repeatSectionTable.out.pdf' ) )
    def testSpreadsheet( self ):
        path = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/spreadsheet.ods' ) )
        pmm = pyMailMerge( path, 'ods' )
        x = [
               { 'token':'repeatrow|invoice', 'value':['1', '2', '3', '4'] },
               { 'token':'total', 'value':['1213.23' ,'531.34', '654.21', '3123.3'] },
               { 'token':'date', 'value':['2011-01-01','2011-01-07','2011-01-03','2011-01-02'] }
        ]
        pmm._process( x )
        pmm.document.refresh()
        pmm.document.saveAs( os.path.join( os.path.dirname( __file__ ), 'docs/spreadsheet.out.xls' ) )
if __name__ == '__main__':
    unittest.main()