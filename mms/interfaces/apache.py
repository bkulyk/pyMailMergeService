import sys
sys.stdout = sys.stderr
import atexit, threading
import cherrypy as http
from mms import init_locale
from mms.interfaces.rest import rest
import mms.config as mms_config
import uuid

class apache( rest ):
    pass

http.config.update( {'environment':'embedded'} )
if http.__version__.startswith( '3.0' ) and http.engine.state == 0:
    http.engine.start( blocking=False )
    atexit.register( http.engine.stop )

config = mms_config.getConfig()
init_locale()
application = http.Application( apache( config ), script_name=None, config=None )