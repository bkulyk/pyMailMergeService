#!/usr/bin/env python
from sys import path
import unittest, os, time
path.append( '..' )
from mms.OfficeDocument import OfficeConnection
from mms.OfficeDocument.WriterDocument import WriterDocument

class mytest():
    def __del__(self):
        print "delete me"

#define unit tests
class testOfficeConnection( unittest.TestCase ):
    path = ''
    def setUp(self):
        self.path = os.path.dirname( os.path.abspath( __file__ ) )
        
    def testIsOfficeRunning(self):
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Libre/Open office should not be running.  Shut it down and try again." )
        
    def testStartAndStopOffice(self):
        #should not be running
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Libre/Open office should not be running.  Shut it down and try again." )
        
        #start and make sure it started
        OfficeConnection.startOffice()
        self.assertTrue( OfficeConnection.isOfficeRunning() )
        
        #stop office and make sure it stopped
        OfficeConnection.stopOffice()
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Something went wrong Libre/Open office should not be running." )
        
    def testOpeningDocumentStartsOffice(self):
        #should not be running
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Libre/Open office should not be running.  Shut it down and try again." )
        
        #opening a document should cause office to start
        infile = "%s/docs/a.odt" % self.path
        w = WriterDocument()
        self.assertTrue( OfficeConnection.isOfficeRunning() )
        
        #stop office and make sure it stopped
        OfficeConnection.stopOffice()
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Something went wrong Libre/Open office should not be running." )
        
    def testOfficeShutsDownAutomatically(self):
         #should not be running
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Libre/Open office should not be running.  Shut it down and try again." )
        
        #start and make sure it started
        OfficeConnection.startOffice()
        self.assertTrue( OfficeConnection.isOfficeRunning() )
        
        #destroying this object should cause office to stop
        del( OfficeConnection.instance )
        
        #should not be running
        self.assertFalse( OfficeConnection.isOfficeRunning(), "Something went wrong Libre/Open office should not be running." )
    
if __name__ == '__main__':
    unittest.main()