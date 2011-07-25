import sys
sys.stdout = sys.stderr
import atexit, threading
import cherrypy as http
from mms.interfaces.rest import rest
from mms import mms
import uuid

class apache( rest ):
    pass

http.config.update( {'environment':'embedded'} )
if http.__version__.startswith( '3.0' ) and http.engine.state == 0:
    http.engine.start( blocking=False )
    atexit.register( http.engine.stop )

config = mms.parseConfig()
application = http.Application( apache( config ), script_name=None, config=None )
