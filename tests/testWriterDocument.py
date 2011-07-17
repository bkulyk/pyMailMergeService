#!/usr/bin/env python
from sys import path, exit
path.append( '..' )
from mms.OfficeDocument import *
from mms.OfficeDocument.WriterDocument import WriterDocument
import unittest, os, uno

class testWriterDocument( unittest.TestCase ):
    path = ''
    def setUp(self):
	self.path = os.path.dirname( os.path.abspath( __file__ ) )
        pass
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

