import cherrypy as http
from mms.interfaces import base
import simplejson as json
class rest( base ):
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
        return base.batchConvert( self, batch, outputType, outputFile )
    @http.expose
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return base.uploadConvert( params, odt, type )
    @http.expose
    def batchpdf( self, batch ):
        return base.batchpdf( self, batch )
    @http.expose
    def pdf( self, params='', odt='' ):
        return base.convert( self, params, odt, 'pdf' )
    @http.expose
    def convert( self, params='', odt='', type='pdf', resave=False, saveExport=False ):
        return base.convert( self, params, odt, type, resave, saveExport )
    @http.expose
    def getTokens( self, odt='', format='json' ):
        return base.getTokens( self, odt, format )
    @http.expose
    def calculator( self, params='', ods='', format='json', outputFile=None ):
        return base.calculator( self, params, ods, format, outputFile )
    @http.expose
    def getNamedRanges( self, ods, format='json' ):
        return base.getNamedRanges( self, ods, format )
    @staticmethod
    def run( options={} ):
        http.config.update( {'server.socket_host':'0.0.0.0', 'server.socket_port':8080} )
        http.quickstart( rest( options ) )
