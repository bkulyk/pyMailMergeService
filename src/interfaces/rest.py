import cherrypy as http
import sys
sys.path.append( ".." )
import interfaces
import simplejson as json
class rest( interfaces.base ):
    @http.expose
    def index(self, test='default'):
        return """
        <html><body>
        <form action='/test' method='post'>
        <h1>Mail Merge Service</h1>
        <p>%s</p>
        </form></body>
        </html>
        """ % test
    @http.expose
    def batchConvert( self, batch, outputType='pdf', outputFile=False ):
        return interfaces.base.batchConvert( self, batch, outputType, outputFile )
    @http.expose
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return interfaces.base.uploadConvert( params, odt, type )
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
    @http.expose
    def calculator( self, params='', ods='', format='josn' ):
        return interfaces.base.calculator( self, params, ods, format )
    @http.expose
    def getNamedRanges( self, ods, format='json' ):
        return interfaces.base.getNamedRanges( self, ods, format )
    @staticmethod
    def run( options={} ):
        http.config.update( {'server.socket_host':'0.0.0.0', 'server.socket_port':8080} )
        http.quickstart( rest( options ) )
