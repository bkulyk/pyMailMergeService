import os.path
from pyMailMerge import pyMailMerge
from SOAPpy import *                #soap interface
import base64                       #to be able to return the output file utf encoded
import tempfile                     #create temp files for writing xml and output
import rest
class soap( rest ):
    documentBase = "tests/docs/"
    def getMethods(self):
        """list the names of the methods that soap is allowed to use"""
        x = []
        for i in dir( self ):
            if i[0] != '_' and i != 'getMethods':
                if callable( getattr( self, i ) ):
                    x.append( i )
        return x
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return rest.uploadConvert( params, base64.b64decode(odt), type )
    def pdf( self, params='', odt='' ):
        return self.convert(params, odt, 'pdf')
    def convert( self, params='', odt='', type='pdf' ):
        try:
            fileName = os.path.abspath( self.documentBase + odt )
            mms = pyMailMerge( fileName )
            return base64.b16decode( mms.convert( params, type ) )
        except:
            number = '?'
            message = "unknown exception"
        return self.__errorXML( number, message )
    def getTokens( self, odt='', format='json' ):
        try:
            path = os.path.dirname( __file__ )
            path = os.path.join( path, self.documentBase, odt )
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
    @staticmethod
    def run():
        host = 'localhost'
        port = 8888
        server = SOAPServer( ( host, port ) )
        soapmms = soap( port=port, host=host )
        namespace = 'com.mailmergeservice'
        #I used to just register the object and let SOAPpy worry about what methods can and can't get 
        #called, but I found that there was no way to get the SOAPContext if I did that.  So I added 
        #a method to pyMailMerge to return all of the public methods automatically and then they are
        #registered one by one.
        for x in soapmms.getMethods():
            server.registerFunction( MethodSig( getattr(soap,x), keywords=0, context=1 ), namespace )
        server.serve_forever()