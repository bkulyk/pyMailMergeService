#!/usr/bin/env python
import os.path, sys
from lib.pyMailMerge import pyMailMerge
import simplejson as json
from interfaces import rest
from daemon import Daemon
import time
#need to extend daemon and implement the run method in order to tell daemon what to do 
class mmsd( Daemon ):
    config = None
    def __init__( self ):
        Daemon.__init__( self, '/tmp/mms.pid', stderr='/tmp/mms.error.log' )
    def run( self ):
        inifile = os.path.join( os.path.abspath( os.path.dirname( __file__ ) ), 'mms.ini' )
#        self.config = ConfigParser.ConfigParser()
#        self.config.read( inifile )
#        return
        rest.rest.run()
#    def get( self, option, default='' ):
#        pass
#parse arguments and take necessary actions
if __name__ == "__main__":
    daemon = mmsd()
    if len( sys.argv ) == 2:
        if sys.argv[1]  == 'start':
            daemon.start()
        elif sys.argv[1] == 'stop':
            daemon.stop()
        elif sys.argv[1] == 'restart':
            daemon.restart()
        #ability simply run without the daemon, useful for debugging
        elif sys.argv[1] == '--no-daemon':
            daemon.run()
        sys.exit(0)
    elif len( sys.argv ) == 1:
        print "usage: %s start|stop|restart|--no-daemon" % sys.argv[0]
        sys.exit( 2 )