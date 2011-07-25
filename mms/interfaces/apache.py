import sys
sys.stdout = sys.stderr
import atexit, threading
import cherrypy as http
from mms.interfaces.rest import rest
from mms import mms
import uuid

class apache( rest ):
#    config = None
#
#    def __init__( self, config ):
#        self.config = config
#
#    @http.expose
#    def index( self ):
#        num = uuid.uuid4()
#        iface = self.config.get( 'mms', 'interface' )
#        return """hello world!<p/>%s<p/>interface: %s<p/><a href='/crash'>crash me</a>""" % ( num, iface )
#
#    @http.expose
#    def crash( self ):
#        import ctypes
#        i = ctypes.c_char( 'a' )
#        j = ctypes.pointer( i )
#        c = 0
#        while True:
#            j[c] = 'a'
#            c += 1
#        j
	pass

http.config.update( {'environment':'embedded'} )
if http.__version__.startswith( '3.0' ) and http.engine.state == 0:
    http.engine.start( blocking=False )
    atexit.register( http.engine.stop )

config = mms.parseConfig()
application = http.Application( apache( config ), script_name=None, config=None )
