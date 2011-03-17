#!/usr/bin/env python
import cherrypy as http
import os.path, sys
from pyMailMerge import pyMailMerge
import simplejson as json
from interfaces import rest
from daemon import Daemon
import time
#need to extend daemon and implement the run method in order to tell daemon what to do 
class mmsd( Daemon ):
    def run( self ):
        rest.rest.run()
#parse arguments and take necessary actions
if __name__ == "__main__":
    stderr = os.path.join( os.path.dirname( __file__ ), 'error.txt' )
    daemon = mmsd( '/tmp/pyMailMergeServiced.pid' ) #, stderr=stderr )
    if len( sys.argv ) == 2:
        if sys.argv[1]  == 'start':
            daemon.start()
        elif sys.argv[1] == 'stop':
            daemon.stop()
        elif sys.argv[1] == 'restart':
            daemon.restart()
        #ability simply run without the daemon, useful for debugging
        elif sys.argv[1] == '--no-daemon':
            rest.rest.run()
        sys.exit(0)
    elif len( sys.argv ) == 1:
        print "usage: %s start|stop|restart" % sys.argv[0]
        sys.exit( 2 )