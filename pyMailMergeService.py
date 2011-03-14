import cherrypy as http
import os.path
from pyMailMerge import pyMailMerge
import simplejson as json
class rest:
    documentBase = "tests/docs/"
    @http.expose
    def index(self, test='default'):
        return """
        <html><body>
        <form action='/test' method='post'>
        <h1>Hello World !</h1>
        <input type='text' value='%s' name='var' />
        <input type='submit' value='Submit' />
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
        except:
            number = "?"
            message = "unknown exception"
        return self.__errorXML( number, message )
    @http.expose
    def getTokens( self, odt='', format='json' ):
        try:
            mms = pyMailMerge( self.documentBase + odt )
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
    http.root = rest()
    http.server.start()