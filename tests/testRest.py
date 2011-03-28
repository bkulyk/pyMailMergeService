#!/usr/bin/env python
from sys import path
path.append( '../src' )
from pyMailMerge import *
import unittest
import os
import urllib2
import urllib
#define unit tests
class testRest( unittest.TestCase ):
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
    def testPDF(self):
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
if __name__ == '__main__':
    unittest.main()