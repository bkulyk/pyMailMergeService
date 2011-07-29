#!/usr/bin/env python
from sys import path
import unittest
path.append( '..' )
from mms.OfficeDocument import OfficeConnection

#define unit tests
class testRest( unittest.TestCase ):
    def testIsOfficeRunning(self):
        x = OfficeConnection.isOfficeRunning()
        self.assertTrue( x )
        
if __name__ == '__main__':
    unittest.main()