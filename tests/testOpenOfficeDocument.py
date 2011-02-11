#!/usr/bin/env python
from sys import path
from sys import exit
path.append( '..' )
from OpenOfficeDocument import *
import unittest
import os
import uno
#define unit tests
class testOpenOfficeDocument( unittest.TestCase ):
    def setUp(self):
        pass
	'''
    def test_searchAndReplaceWithDocument( self ):
	ood = OpenOfficeDocument()
	ood.open( '/usr/share/pyMailMergeService/tests/docs/find_replace.odt' )	
	replace = ood.oodocument.createReplaceDescriptor()
	replace.setSearchString( "search" )
	replace.setReplaceString( "replace" )
	ood.oodocument.replaceAll( replace )
	ood.saveAs( '/usr/share/pyMailMergeService/tests/docs/find_replaced.pdf' )
	'''
    def test_insertDocument( self ):
	ood = OpenOfficeDocument()
	ood.open( '/usr/share/pyMailMergeService/tests/docs/find_replace.odt' )
	
	cursor = ood.oodocument.Text.createTextCursor()
	
	replace = ood.oodocument.createReplaceDescriptor()
	replace.setSearchString( "search" )
	replace.setReplaceString( "replace" )
	ood.oodocument.replaceAll( replace )

	ood.oodocument.Text.insertString( cursor, 'inserted', 0 )

	#print ood.oodocument
	properties = []
	#properties.append( OpenOfficeDocument._makeProperty( "Hidden", True ) )
	properties = tuple( properties )
	cursor.insertDocumentFromURL( uno.systemPathToFileUrl("/usr/share/pyMailMergeService/tests/docs/insert_doc.odt"), properties )
	ood.saveAs( "/usr/share/pyMailMergeService/tests/docs/inserted_doc.pdf" )
	ood.close()

if __name__ == '__main__':
    unittest.main()
	
'''	def test_RandomStuff(self):
        """This test doen't do anything except serve as a place holder for some code, i'm not
        yet willing to delete."""
        filterFactory = uno.createUnoStruct( 'com.sun.star.document.FilterFactory' )
        x = filterFactory.getElementNames()
        host = 'localhost'
        port = 8100
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
        context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
        """This code retrieves a list of all the import/export filters supported by OOo"""
        desktop = context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', context )
        doc = desktop.loadComponentFromURL( uno.systemPathToFileUrl( os.path.abspath( "docs/if_section_simple.odt" ) ), "_blank", 0, () )
        filterFactory = local.ServiceManager.createInstanceWithContext( 'com.sun.star.document.FilterFactory', context )
        for x in filterFactory.getElementNames():
            for y in filterFactory.getByName( x ):
                #print '%s ============= %s' % (y.Name, y.Value)
                pass
            #print ''
	'''
