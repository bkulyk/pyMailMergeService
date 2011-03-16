import cherrypy as http
import os.path, sys
from pyMailMerge import pyMailMerge
import simplejson as json
from interfaces import rest
from daemon import Daemon
import time

class pyMailMergeServiced( Daemon ):
    def run( self ):
        rest.rest.run()

if __name__ == "__main__":
    stderr = os.path.join( os.path.dirname( __file__ ), 'error.txt' )
    daemon = pyMailMergeServiced( '/tmp/pyMailMergeServiced.pid', stderr=stderr )
    if len( sys.argv ) == 2:
        if sys.argv[1]  == 'start':
            daemon.start()
        elif sys.argv[1] == 'stop':
            daemon.stop()
        elif sys.argv[1] == 'restart':
            daemon.restart()
        sys.exit(0)
    elif len( sys.argv ) == 1:
        print "usage: %s start|stop|restart" % sys.argv[0]
        sys.exit( 2 )