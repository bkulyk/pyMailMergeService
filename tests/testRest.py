#!/usr/bin/env python
from sys import path
import unittest
import os
import urllib2
import urllib
import subprocess, time
import signal
import simplejson as json
from lxml import etree
#define unit tests
class testRest( unittest.TestCase ):
    process = None
    startRest = False
    devnull = None
    def setUp(self):
        if self.startRest:
            self.devnull = open('/dev/null')
            path = os.path.dirname( os.path.abspath( os.path.join( os.path.dirname( __file__ ) ) ) )
            path = os.path.join( path, 'src/mmsd.py' )

            self.process = subprocess.Popen( [ '/usr/bin/env', 'python', path, '--no-daemon' ], stdout=self.devnull, stderr=self.devnull )
            time.sleep( 1 )
        
    def tearDown(self):
        if self.startRest:
            self.process.kill()
            self.devnull.close()
    
    xml = """<?xml version="1.0" encoding="UTF-8"?>
        <tokens>
            <token>
                <name>company::name</name>
                <value>REST Company</value>
            </token>
            <token>
                <name>company::address1</name>
                <value>12 whatever st.</value>
            </token>
            <token>
                <name>company::city</name>
                <value>Winnipeg</value>
            </token>
            <token>
                <name>company::prov</name>
                <value>MB</value>
            </token>
            <token>
                <name>company::phone</name>
                <value>204-555-1234</value>
            </token>
            <token>
                <name>company::postalcode</name>
                <value>R3M 1D3</value>
            </token>
            
            <token>
                <name>client::name</name>
                <value>Some Guy</value>
            </token>
            <token>
                <name>client::company</name>
                <value>Client Company</value>
            </token>
            <token>
                <name>client::prov</name>
                <value>ON</value>
            </token>
            <token>
                <name>client::prov</name>
                <value>AB</value>
            </token>
            <token>
                <name>client::city</name>
                <value>Calgary</value>
            </token>
            <token>
                <name>client::postalcode</name>
                <value>T2S 0E1</value>
            </token>
            <token>
                <name>repeatrow|product::desc</name>
                <value>aasdf asdhfasdf heradfasd hwe</value>
            </token>
            <token>
                <name>product::rate</name>
                <value>12.12</value>
            </token>
            <token>
                <name>product::qty</name>
                <value>1</value>
            </token>
            <token>
                <name>product::total</name>
                <value>12.12</value>
            </token>
            <token>
                <name>product::rate</name>
                <value>12.12</value>
            </token>
            <token>
                <name>if|paid</name>
                <value>1</value>
            </token>
            <token>
                <name>if|notpaid</name>
                <value>0</value>
            </token>
            <token>
                <name>paid::date</name>
                <value>Jan 01, 2011</value>
            </token>
            <token>
                <name>payment::due</name>
                <value>Jan 01, 2011</value>
            </token>
            <token>
                <name>paid</name>
                <value>PAID</value>
            </token>
            <token>
                <name>html|notes</name>
                <value><![CDATA[<div style="font-family: Arial;"><h3>Terms:</h3><ol><li>Payment due in <strong>30</strong> days.</li><li>No refunds</li></ol></div>]]></value>
            </token>
        </tokens>"""
    '''def testPDF(self):
        xml = self.xml
        try:
            data = { 'params':xml, 'odt':"invoice.odt" }
            url = "http://localhost:8080/pdf"
            x = urllib2.urlopen( url, urllib.urlencode( data ) )
            file = open( os.path.join( os.path.dirname( __file__ ), 'docs/rest.out.pdf' ), 'w' )
            file.write( x.read() )
            file.close()
            allGood = True
        except:
            allGood = False
        self.assertTrue( allGood )
    def testUploadConvert(self):
        xml = self.xml
        odtFilename = os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'docs/invoice.odt' ) )
        file = open( odtFilename, 'r' )
        odt = file.read()
        file.close()
        try:
            data = { 'params':xml, 'odt':odt }
            url = "http://localhost:8080/uploadConvert"
            x = urllib2.urlopen( url, urllib.urlencode( data ) )
            file = open( os.path.join( os.path.dirname( __file__ ), 'docs/rest_upload_convert.out.pdf' ), 'w' )
            file.write( x.read() )
            file.close()
            allGood = True
        except:
            allGood = False
        self.assertTrue( allGood )
    '''
    def testCalculator(self):
        f = open( os.path.abspath( os.path.join( os.path.dirname( __file__ ), 'fixtures/calculator.xml' ) ) )
        
        data = { 'params':f.read(), 'ods':'../tests/docs/calculator.ods' }
        response = urllib2.urlopen( "http://localhost:8080/calculator", urllib.urlencode( data ) )
        body = response.read()
        
        expected = { 'totals':[ '1514', '1668', '910' ],
                    'test':[ ['a','b'],['c','d'],['e','f'],['g','h'] ],
                    'results':[ '151.4', '333.6', '455' ],
                    'more':[ '124', '548', '464' ]
                   }
        results = json.loads( body )
        self.assertEquals( expected, results )
    
    def testGetNamedRanges(self):
        data = { 'ods':'../tests/docs/calculator.ods' }
        response = urllib2.urlopen( "http://localhost:8080/getNamedRanges", urllib.urlencode( data ) )
        body = response.read()
        
        expected = [ 'first', 'second', 'third', 'totals', 'factors', 'results', 'test', 'more' ]
        results = json.loads( body )
        
        for x in expected:
            self.assertTrue( x in results, "%s is missing from results" % x )
        
    def testGetNamedRanges_XML(self):
        data = { 'ods':'../tests/docs/calculator.ods', 'format':'xml' }
        response = urllib2.urlopen( "http://localhost:8080/getNamedRanges", urllib.urlencode( data ) )
        xml = response.read()
        
        expected = [ 'first', 'second', 'third', 'totals', 'factors', 'results', 'test', 'more' ]
        
        xml = etree.XML( xml )
        elements = xml.xpath( "//namedranges/namedrange" )
        
        results = []
        for x in elements:
            results.append( x.text ) 
            self.assertTrue( x.text in expected, 'was not expecting "%s" in results' % x )
        
        for x in expected:
            self.assertTrue( x in results, 'was expecting "%s" in results' % x )
    
    def testBatchConvert(self):

        data = [ [ "../tests/docs/batchConvert_part1.odt", '' ], 
                 [ "../tests/docs/batchConvert_part2.odt", '' ],
                 [ "../tests/docs/batchConvert_part3.odt", '' ] ]
        
        data = { 'batch':json.dumps( data ), 'outputType':'odt' }
        
        response = urllib2.urlopen( "http://localhost:8080/batchConvert", urllib.urlencode( data ) )
        
        f = open( 'docs/batchConvert.out.odt', 'w' )
        f.write( response.read() )
        f.close()
        
if __name__ == '__main__':
    unittest.main()