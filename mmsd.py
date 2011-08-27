#!/usr/bin/env python
import os.path, sys, time, getopt
import simplejson as json
import mms
from mms.daemon import Daemon

#need to extend daemon and implement the run method in order to tell daemon what to do 
class mmsd( Daemon ):
    config = None
    options = {}
    def __init__( self ):
        Daemon.__init__( self, '/tmp/mms.pid', stderr='/tmp/mms.error.log' )
    def run( self ):
        mms.run_mms()

#parse arguments and take necessary actions
if __name__ == "__main__":
    command = ''

    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument( 'command', type=str, choices=( 'start','stop','restart','nodaemon' ), help="start, stop, restart the Mail Merge Service. Or start the service without the daemon." )
    
    args = parser.parse_args()
    
    if args.command == 'nodaemon':
        args.command = 'run'
   
    if command == 'no-daemon':
        mms.run_mms()
    else:
        daemon = mmsd()
        #run command
        getattr(daemon, args.command)()
