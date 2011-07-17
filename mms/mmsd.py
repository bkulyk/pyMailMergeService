#!/usr/bin/env python
import os.path, sys, time, getopt
import simplejson as json
from mms import pyMailMerge
from mms.interfaces import rest
from mms.daemon import Daemon
#need to extend daemon and implement the run method in order to tell daemon what to do 
class mmsd( Daemon ):
    config = None
    options = {}
    def __init__( self ):
        Daemon.__init__( self, '/tmp/mms.pid', stderr='/tmp/mms.error.log' )
    def run( self ):
        inifile = os.path.join( os.path.abspath( os.path.dirname( __file__ ) ), 'mms.ini' )
#        self.config = ConfigParser.ConfigParser()
#        self.config.read( inifile )
#        return
        print self.options
        rest.rest.run( self.options )
#    def get( self, option, default='' ):
#        pass

def usage():
    print "Add me"

#parse arguments and take necessary actions
if __name__ == "__main__":
    command = 'run'
    stubsDir = ''
    outputDir = ''
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hm:", ["help", "no-daemon",
            "output-dir=", "stubs-dir=", "mode="])
    except getopt.GetoptError, err:
        print str(err)
        usage()
        sys.exit(2)

    for o, a in opts:
        if o in ("-h", "--help"):
            usage()
            sys.exit()
        elif o in ("--no-daemon"):
            command = 'run'
        elif o in ("-m", "--mode"):
            if command != 'run':
                command = a
        elif o in ("--output-dir"):
            outputDir = a
        elif o in ("--stubs-dir"):
            stubsDir = a
        else:
            usage()
            sys.exit()

    daemon = mmsd()
    if outputDir != '':
        daemon.options['outputDir'] = outputDir
    if stubsDir != '':
        daemon.options['stubsDir'] = stubsDir
    getattr(daemon, command)()
