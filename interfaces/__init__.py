#!/usr/bin/env python
import os, sys
sys.path.append( ".." )
from pyMailMerge import pyMailMerge
import simplejson as json
import tempfile
class base( object ):
    documentBase = "../tests/docs/"
    def uploadConvert( self, params='', odt='', type='pdf' ):
        #write temporary file
        file, fileName = tempfile.mkstemp( suffix="_pyMMS.odt" )
        os.close( file )
        file = open( fileName, 'w' )
        file.write( odt )
        file.close()
        #do conversion
        try:
            mms = pyMailMerge( fileName )
            content = mms.convert( params, type )
            os.unlink( fileName )
            return content
        except:
            number = '?'
            message = "unknown exception"
        return self.__errorXML( number, message )
    def pdf( self, params='', odt='' ):
        return self.convert(params, odt, 'pdf')
    def convert( self, params='', odt='', type='pdf' ):
        try:
            fileName = os.path.abspath( self.documentBase + odt )
            mms = pyMailMerge( fileName )
            return mms.convert( params, type )
        except:
            number = '?'
            message = "unknown exception"
        return self.__errorXML( number, message )
    def getTokens( self, odt='', format='json' ):
        try:
            path = os.path.dirname( __file__ )
            path = os.path.join( path, self.documentBase, odt )
            print os.path.abspath( path )
            mms = pyMailMerge( os.path.abspath( path ) )
            tokens = mms.getTokens()
            if format=='xml':
                xml = """<?xml version="1.0" encoding="UTF-8"?><tokens>"""
                for x in tokens:
                    xml += "<token>%s</token>" % x
                xml += "</tokens>"
                return xml
            else:
                return json.dumps( tokens )
        except:
            return self.__errorXML( '?', 'could not get tokens' )
    def __errorXML( self, number, message ):
        return """
        <?xml version="1.0" encoding="UTF-8"?>
        <errors>
            <error>
                <number>%s</number>
                <message>%s</message>
            </error>
        </errors>""" % (number, message)