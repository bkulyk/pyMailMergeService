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
    def joinDocuments( self, odt, fileNames, addPageNumbers=False, appendToExistingDoc=False ):
        return interfaces.base.joinDocuments( self, odt, fileNames, addPageNumbers, appendToExistingDoc )
    @http.expose
    def copyDocument( self, target, source ):
        return interfaces.base.copyDocument( self, target, source )
    @http.expose
    def batchpdf( self, batch ):
        return interfaces.base.batchpdf( self, batch )
    @http.expose
    def pdf( self, params='', odt='' ):
        return interfaces.base.convert( self, params, odt, 'pdf' )
    @http.expose
    def convert( self, params='', odt='', type='pdf', resave=False, saveExport=False ):
        return interfaces.base.convert( self, params, odt, type, resave, saveExport )
    @http.expose
    def getTokens( self, odt='', format='json' ):
        return interfaces.base.getTokens( self, odt, format )
    @staticmethod
    def run( options={} ):
        http.config.update( {'server.socket_host':'0.0.0.0', 'server.socket_port':8080} )
        http.quickstart( rest( options ) )
