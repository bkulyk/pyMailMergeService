import cherrypy as http
import sys
sys.path.append( ".." )
import interfaces
class rest( interfaces.base ):
#    @http.expose
#    def index(self, test='default'):
#        return """
#        <html><body>
#        <form action='/test' method='post'>
#        <h1>%s</h1>
#        </form></body>
#        </html>
#        """ % test
    @http.expose
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return interfaces.base.uploadConvert( params, odt, type )
    @http.expose
    def joinDocuments( self, odt, fileNames ):
        return interfaces.base.joinDocuments( self, odt, fileNames )
    @http.expose
    def batchpdf( self, batch ):
        return interfaces.base.batchpdf( self, batch )
    @http.expose
    def pdf( self, params='', odt='' ):
        return interfaces.base.convert( self, params, odt, 'pdf' )
    @http.expose
    def convert( self, params='', odt='', type='pdf' ):
        return interfaces.base.convert( self, params, odt, type )
    @http.expose
    def getTokens( self, odt='', format='json' ):
        return interfaces.base.getTokens( self, odt, format )
    @staticmethod
    def run():
	http.config.update( {'server.socket_host':'0.0.0.0', 'server.socket_port':80} )
        http.quickstart( rest() )
