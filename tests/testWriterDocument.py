#!/usr/bin/env python
from sys import path
from sys import exit
path.append( '..' )
path.append( '../src' )
path.append( '../src/OfficeDocument' )
from OfficeDocument.WriterDocument import *
import unittest, os, uno
#define unit tests
class testWriterDocument( unittest.TestCase ):
    path = ''
    def setUp(self):
	self.path = os.path.dirname( os.path.abspath( __file__ ) )
        pass
    def test_fromBase26(self):
        self.assertEquals( 1,  B26.fromBase26( 'A' ) )
        self.assertEquals( 27, B26.fromBase26( 'AA' ) )
        self.assertEquals( 26, B26.fromBase26( 'Z' ) )
        self.assertEquals( 2,  B26.fromBase26( 'B' ) )
        self.assertEquals( 53, B26.fromBase26( 'BA' ) )
        self.assertEquals( 54, B26.fromBase26( 'BB' ) )
    def test_duplicateRow(self):
        ood = WriterDocument()
        ood.open( '%s/docs/repeat_row_repeat_column.odt' % self.path )
        ood.duplicateRow( "~a2~", True )
        ood.saveAs( '%s/docs/repeat_row_repeat_column.pdf' % self.path )
        ood.close()
    def test_duplicateColumn(self):
        ood = WriterDocument()
        ood.open( '%s/docs/repeat_row_repeat_column.odt' % self.path )
        ood.duplicateColumn( "~a1~", True )
        ood.saveAs( '%s/docs/repeat_row_repeat_column.2.pdf' % self.path )
        ood.close()
    def test_searchAndDuplicate( self ):
        ood = WriterDocument()
        ood.open( "%s/docs/duplicate_section.odt" % self.path )
        ood.searchAndDuplicate( "~start~", '~end~', 3, True )
        ood.saveAs( "%s/docs/duplicate_section.pdf" % self.path )
        ood.close()    
    def test_searchAndCursor( self ):
    	ood = WriterDocument()
    	ood.open( "%s/docs/find_replace.odt" % self.path )
    	search = ood.oodocument.createSearchDescriptor()
    	search.setSearchString( 'search' )
    	result = ood.oodocument.findFirst( search )
    	path = uno.systemPathToFileUrl( "%s/docs/insertme.html" % self.path )
    	result.insertDocumentFromURL( path, tuple() )
    	ood.saveAs( "%s/docs/docInTheMiddle.pdf" % self.path )
    	ood.close()
    def test_searchAndReplaceWithDocument( self ):
    	ood = WriterDocument()
    	ood.open( '%s/docs/find_replace.odt' % self.path )
    	replace = ood.oodocument.createReplaceDescriptor()
    	replace.setSearchString( "search" )
    	replace.setReplaceString( "replace" )
    	ood.oodocument.replaceAll( replace )
    	ood.saveAs( '%s/docs/find_replaced.pdf' % self.path )
        ood.close()
    def test_insertDocument( self ):
    	ood = WriterDocument()
        ood.open( '%s/docs/find_replace.odt' % self.path )
    	cursor = ood.oodocument.Text.createTextCursor()
    	replace = ood.oodocument.createReplaceDescriptor()
    	replace.setSearchString( "search" )
    	replace.setReplaceString( "replace" )
    	ood.oodocument.replaceAll( replace )
    	ood.oodocument.Text.insertString( cursor, 'inserted', 0 )
    	properties = []
    	properties = tuple( properties )
    	cursor.insertDocumentFromURL( uno.systemPathToFileUrl("%s/docs/insert_doc.odt" % self.path), properties )
    	ood.saveAs( "%s/docs/inserted_doc.pdf" % self.path )
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
