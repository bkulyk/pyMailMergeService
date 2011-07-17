from SOAPpy import *
import base64, sys
import interfaces
class soaphelper:
    exposedMethods = {}
    @staticmethod
    def expose( func=None ):
        soaphelper.exposedMethods[ func.__name__ ] = func
class soap( interfaces.base ):
    @soaphelper.expose
    def uploadConvert( self, params='', odt='', type='pdf' ):
        return interfaces.base.uploadConvert( self, params, base64.b64decode(odt), type )
    @soaphelper.expose
    def batchpdf( self, batch ):
        return base64.b64encode( interfaces.base.batchpdf( self, batch ) )
    @soaphelper.expose
    def pdf( self, params='', odt='' ):
        return self.convert( params, odt, 'pdf' )
    @soaphelper.expose
    def convert( self, params='', odt='', type='pdf' ):
        return base64.b64encode( self, interfaces.base.convert( self, params, odt, type ) )
    @soaphelper.expose
    def getTokens( self, odt='', format='json' ):
        return interfaces.base.getTokens( self, odt, format )
    @staticmethod
    def run():
        host = 'localhost'
        port = 8181
        server = SOAPServer( ( host, port ) )
        namespace = 'com.mailmergeservice'
        for k in soaphelper.exposedMethods:
            server.registerFunction( MethodSig( soaphelper.exposedMethods[k], keywords=0, context=1 ), namespace )
        server.serve_forever()
if __name__ == '__main__':
    soap.run()