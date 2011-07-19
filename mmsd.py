#!/usr/bin/env python
import os.path, sys, time, getopt
import simplejson as json
from mms import mms
#from mms.interfaces import rest
#from mms.interfaces.rest import rest
from mms.daemon import Daemon
#need to extend daemon and implement the run method in order to tell daemon what to do 
class mmsd( Daemon ):
    config = None
    options = {}
    def __init__( self ):
        Daemon.__init__( self, '/tmp/mms.pid', stderr='/tmp/mms.error.log' )
    def run( self ):
        i = mms()

#parse arguments and take necessary actions
if __name__ == "__main__":
    command = ''

    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument( 'command', type=str, choices=( 'start','stop','restart', 'no-daemon' ), help="start, stop, restart the Mail Merge Service. Or start the service without the daemon." )
    parser.add_argument( '--stubs-dir', type=str, help='Folder to source document stubs out of.  Use absolute path.' )
    parser.add_argument( '--output-dir', type=str, help='Path to output documents for operations where output is required.' )
    
    args = parser.parse_args()
    
    if args.command == 'no-daemon':
        args.command = 'run'
    
    daemon = mmsd()
    
    if args.output_dir is not None :
        daemon.options['outputDir'] = args.output_dir
    if args.stubs_dir is not None:
        daemon.options['stubsDir'] = args.stubs_dir
        
    #run command    
    getattr(daemon, args.command)()
