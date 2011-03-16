import cherrypy as http
import os.path
from pyMailMerge import pyMailMerge
import simplejson as json
class rest:
    documentBase = "../tests/docs/"
    @http.expose
    def index(self, test='default'):
        return """
        <html><body>
        <form action='/test' method='post'>
        <h1>%s</h1>
        </form></body>
        </html>
        """ % test
    @http.expose
    def pdf( self, params='', odt='' ):
        return self.convert(params, odt, 'pdf')
    @http.expose
    def convert( self, params='', odt='', type='pdf' ):
        try:
            mms = pyMailMerge( self.documentBase + odt )
            return mms.convert( params, 'pdf' )
        except Exception, e:
            message = "unknown exception"
        return self.__errorXML( e.errno, e.strerror )
    @http.expose
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
        http.quickstart( rest() )