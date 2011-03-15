import cherrypy as http
import os.path
import tempfile
from pyMailMerge import pyMailMerge
import simplejson as json
class rest:
    documentBase = "tests/docs/"
    @http.expose
    def index(self, test='default'):
        return """
        <html><body>
        <h1>Hello World !</h1>
        <p>%s</p>
        </body>
        </html>
        """ % test
    @http.expose
    def uploadConvert( self, odt, params=[], type='pdf' ):
        try:
            file, fileName = tempfile.mkstemp( suffix="_pyMMS.odt" )
            file.write( odt )
            os.close( file )
        except:
            return self.__errorXML( '?', 'failed to write temporary odt file' )
        try:
            mms = pyMailMerge( fileName )
            return mms.convert( params, type )
        except:
            number = "?"
            message = "unknown exception"
        return self.__errorXML( number, message )
    @http.expose
    def pdf( self, params='', odt='' ):
        return self.convert(params, odt, 'pdf')
    @http.expose
    def convert( self, params='', odt='', type='pdf' ):
        try:
            mms = pyMailMerge( self.documentBase + odt )
            return mms.convert( params, type )
        except:
            number = "?"
            message = "unknown exception"
        return self.__errorXML( number, message )
    @http.expose
    def getTokens( self, odt='', format='xml' ):
        try:
            mms = pyMailMerge( self.documentBase + odt )
            tokens = mms.getTokens()
            if format=='xml':
                #convert params to xml
                xml = """<?xml version="1.0" encoding="UTF-8"?><tokens>"""
                for x in tokens:
                    xml += "<token>%s</token>" % x
                xml += "</tokens>"
                return xml
            else:
                #return params as json encoded
                return json.dumps( tokens )
        except:
            number = "?"
            message = "unknown exception"
        return self.__errorXML( number, message )
    def __errorXML( self, number, message ):
        return """
        <?xml version="1.0" encoding="UTF-8"?>
        <errors>
            <error>
                <number>%s</number>
                <message>%s</message>
            </error>
        </errors>""" % (number, message)
if __name__ == '__main__':
    http.quickstart( rest() )