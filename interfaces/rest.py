import cherrypy as http
import sys
sys.path.append( ".." )
import interfaces
class rest( interfaces.base ):
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
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return interfaces.base.uploadConvert( params, odt, type )
    @http.expose
    def pdf( self, params='', odt='' ):
        return interfaces.base.convert( params, odt, 'pdf' )
    @http.expose
    def convert( self, params='', odt='', type='pdf' ):
        return interfaces.base.convert( params, odt, type )
    @http.expose
    def getTokens( self, odt='', format='json' ):
        return interfaces.base.getTokens( self, odt, format )
    @staticmethod
    def run():
        http.quickstart( rest() )