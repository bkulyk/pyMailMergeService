#!/usr/bin/env python
from OpenOfficeDocument import *
from modifiers import *
from lxml import etree              #for parsing xml
import re                           #regular expressions
import operator                     #using for sorting the params
import os                           #removing the temp files
import tempfile                     #create temp files for writing xml and output
class pyMailMerge:
    document = None
    def __init__(self, odt):
        self.document = OpenOfficeDocument()
        self.document.open( odt )
    def __del__(self):
        self.document.close()
    def getTokens(self):
        pass
    def batchpdf(self, odt):
        pass
    def pdf(self, odt, params ):
        pass
    def convert(self, obj):
        pass
    def _process(self, params):
        params = pyMailMerge._sortParams(params)
        for param in params:
            #each module has been set as a property of the modifiers class
            mod = getattr( modifiers, "mod_%s" % param['modifier'] )
            #so, all the modifiers have been loaded as class names called __init__ within their 
            #respective modules, and each modifier should have a method called process
            try:
                mod.__init__().process( self.document, param )
            except:
                pass
    def _getTempFile( self, extension='' ):
        '''
        This is the replacement for os.tmpnam, it should not be suceptable to the
        same symlink injection problems.  I wrapped this simple command in a function 
        because I don't want to use the file handle it creates. There maybe a more effcient
        way to get a tmp file name, but I havn't found one yet, so I'm using this.
        @param extension:  the file extension that the temp file should have.
        @return: str the temporary filename  
        '''
        file = tempfile.mkstemp( suffix="_pyMMS%s" % extension )
        #I am not going to use this file handle because I found in my research,
        #that I need to open it with the codec module to ensure it's open as utf-8 
        #file instead of whatever the default is
        os.close( file[0] )
        return file[1]
    @staticmethod
    def _readParamsFromXML(xml):
        xml = etree.XML( xml )
        tokens = xml.xpath( "//tokens/token" )
        params = []
        for x in tokens:
            tokenname = x.findall( "name" )[0].text
            tokenvalues = x.findall( "value" )
            if len( tokenvalues ) == 1:
                values = tokenvalues[0].text
            else:
                values = []
                for x in tokenvalues:
                    values.append( x.text )
            params.append( { 'token':tokenname, 'value':values } )
        return params
    @staticmethod
    def _sortParams(params):
        #break up by token, so they can be sorted later 
        sorted = { 'None':[] }
        for x in modifiers.modifierOrder:
            sorted[ x['name'] ] = []
        #now place each key in the right spot in sorted
        for x in params:
            result = re.match( r"(?P<modifier>^[A-Za-z]+\|)*(?P<token>[A-Za-z\:\|]+)", x['token'] )
            mod = "%s" % result.group( 'modifier' )
            mod = mod.replace( '|', '' )
            #add the modifier to the dictionary, this will save some effor later
            x['modifier'] = mod
            sorted[ mod ].append( x )
        #so far they are not actually sorted, just divided up, now lets put them all back into one list in the right order
        all = []
        for x in modifiers.modifierOrder:
            all.extend( sorted[ x['name'] ] )
        all.extend( sorted[ 'None' ] )
        return all            