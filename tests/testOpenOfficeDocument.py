#!/usr/bin/env python
from sys import path
path.append( '..' )
from OpenOfficeDocument import *
import unittest
import os
import uno
#define unit tests
class testOpenOfficeDocument( unittest.TestCase ):
    def setUp(self):
        pass
    def test_RandomStuff(self):
#        filterFactory = uno.createUnoStruct( 'com.sun.star.document.FilterFactory' )
#        x = filterFactory.getElementNames()
        host = 'localhost'
        port = 8100
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
        context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
        desktop = context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', context )
        doc = desktop.loadComponentFromURL( uno.systemPathToFileUrl( os.path.abspath( "docs/if_section_simple.odt" ) ), "_blank", 0, () )
        
        filterFactory = local.ServiceManager.createInstance( 'com.sun.star.document.FilterFactory' )
#        print filterFactory
        #print doc
        #ff = doc.FilterFactory
        
if __name__ == '__main__':
    unittest.main()