#!/usr/bin/env python
import uno, os, re
from OfficeDocument import OfficeDocument
from sys import path
#from com.sun.star.beans import PropertyValue
#from com.sun.star.io import IOException
from com.sun.star.style.BreakType import PAGE_BEFORE, PAGE_AFTER
from lib.B26 import B26
class WriterDocument( OfficeDocument ):
    def re_match( self, pattern ):
        search = self.oodocument.createSearchDescriptor()
        search.setSearchString( pattern )
        search.SearchRegularExpression = True
        items = self.oodocument.findAll( search )
        list = []
        for i in xrange( items.getCount() ):
            x = items.getByIndex( i )
            list.append( x.getString() )
        return list
    def addDocumentToEnd( self, documentUrl, pageBreak=True ):
        cursor = self.oodocument.Text.createTextCursor()
        if cursor:
            cursor.gotoEnd( False )
            if pageBreak:
                cursor.BreakType = PAGE_BEFORE
            print documentUrl
            cursor.insertDocumentFromURL( uno.systemPathToFileUrl( os.path.abspath( documentUrl ) ), () )
            return True
        return False
    def searchAndReplaceWithDocument( self, phrase, documentPath, regex=False ):
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_1_1_Editing_Text
    	#cursor = self.oodocument.Text.createTextCursor()
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_3_3_Search_and_Replace
    	search = self.oodocument.createSearchDescriptor()
    	search.setSearchString( phrase )
    	search.SearchRegularExpression = regex
    	result = self.oodocument.findFirst( search )
    	path = uno.systemPathToFileUrl( os.path.abspath( documentPath ) )
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_1_5_Inserting_Text_Files
    	return result.insertDocumentFromURL( path, tuple() )
    def drawSearchAndReplace( self, phrase, replacement ):
        drawPage = self.oodocument.getDrawPage()
        e = drawPage.createEnumeration()
        count = 0
        while e.hasMoreElements():
            x = e.nextElement()
            if x.getString() == phrase:
                x.setString( replacement )
                count += 1
        return count
    def searchAndReplace( self, phrase, replacement, regex=False ):
        replace = self.oodocument.createReplaceDescriptor()
        replace.setSearchString( phrase )
        replace.setReplaceString( replacement )
        replace.SearchRegularExpression = regex
        count = self.oodocument.replaceAll( replace )
        return count
    def searchAndReplaceFirst( self, phrase, replacement, regex=False ):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        if cursor is not None:
            cursor.Text.setString( replacement )
            return 1
        else:
            return 0
    def searchAndDelete( self, phrase, regex=False ):
        self.searchAndReplace( phrase, '', regex )  
    def _getCursorForStartPhrase(self, startPhrase, regex=False):
        #find position of start phrase
        search = self.oodocument.createSearchDescriptor()
        search.setSearchString( startPhrase )
        search.SearchRegularExpression = regex
        result = self.oodocument.findFirst( search )
        return result
    def _getCursorForStartAndEndPhrases( self, startPhrase, endPhrase, regex=False ):
        '''@todo replace with _getCursorForStartPhrase x2 and test'''
        #find position of start phrase
        search = self.oodocument.createSearchDescriptor()
        search.setSearchString( startPhrase )
        search.SearchRegularExpression = regex
        result = self.oodocument.findFirst( search )
        #find position of end phrase
        search2 = self.oodocument.createSearchDescriptor()
        search2.setSearchString( endPhrase )
        search2.SearchRegularExpression = regex
        result2 = self.oodocument.findFirst( search2 )
        try:
            #create new cursor at start of first phrase, expand to second, and return the cursor range
            cursor = self.oodocument.Text.createTextCursorByRange( result.getStart() )
            cursor.gotoRange( result2.getEnd(), True )
            return cursor
        except:
            return None
    def searchAndRemoveSection( self, startPhrase, endPhrase, regex=False ):
        cursor = self._getCursorForStartAndEndPhrases(startPhrase, endPhrase, regex)
        self.oodocument.Text.insertString( cursor, '', True )
    def searchAndDuplicateInTable( self, startPhrase, endPhrase, count, regex=False ):
#        try:
        #get the start and end cells
        c = self._getCursorForStartPhrase( startPhrase, regex )
        enum = c.createEnumeration()
        start = None
        end = None
        if enum.hasMoreElements():
            e = enum.nextElement()
            cellNames = e.getCellNames()
            for x in cellNames:
                cell = e.getCellByName( x )
                text = cell.Text.getString()
                matches = re.match( "\w*%s\w*" % re.escape( startPhrase ), text )
                if matches is not None:
                    start = x
                matches = re.match( "\w*%s\w*" % re.escape( endPhrase ), text )
                if matches is not None:
                    end = x
        #now that we have the start and end cells, copy the content, duplicate and paste content
        if start is not None and end is not None:
            #print "%s - %s" % ( start, end )
            startRow,dummy = self._convertCellNameToCellPositions( start )
            endRow,dummy = self._convertCellNameToCellPositions( end )
            numRowsToAdd = int(endRow) - int(startRow) + 1
            #controller gets the view cursor
            controller = self.oodocument.getCurrentController()
            cellCursor = e.createCursorByCellName( start )
            i = 0
            cells = [] #hold the copied content to paste later 
            while cellCursor is not None:
                if i != 0:
                    """in openoffice/libreoffice if you put your cursor in a cell and 
                    move to the right if you are at the last cell in a row, it will move 
                    the cursor to the first cell of the next row"""
                    cellCursor.goRight( 1, False )
                #copy this cell
                viewCursor = controller.getViewCursor()
                currentCell = e.getCellByName( cellCursor.getRangeName() )
                viewCursor.gotoRange( currentCell.Text, False )
                txt = controller.getTransferable()
                cells.append( txt )
                if cellCursor.getRangeName() == end:
                    break;
                i += 1
            #insert rows into table
            if numRowsToAdd > 1:
                for i in xrange( count ):
                    rows = e.getRows()
                    rows.insertByIndex( endRow+1, numRowsToAdd )
            #paste content into new cells
            for i in xrange( count ):
                for x in xrange( len( cells ) ):
                    cellCursor.goRight( 1, False )
                    currentCell = e.getCellByName( cellCursor.getRangeName() )
                    currentCell.setString( ' ' )
                    viewCursor.gotoRange( currentCell.Text, False )
                    controller.insertTransferable( cells[x] )
#        except:
#            pass
    def searchAndDuplicate( self, startPhrase, endPhrase, count, regex=False ):
        i = 0 #track number of successful duplications
        try:
            cursor = self._getCursorForStartAndEndPhrases( startPhrase, endPhrase, regex )
            if cursor is not None:
                cursor2 = self.oodocument.Text.createTextCursorByRange( cursor.getEnd() )
                if cursor2 is not None:
                    #copy once...
                    controller = self.oodocument.getCurrentController()
                    viewCursor = controller.getViewCursor()
                    viewCursor.gotoRange( cursor, False )
                    txt = controller.getTransferable()
                    #... paste multiple times
                    for x in xrange( count ):
                        #look ma, no more AutoText
                        viewCursor.gotoRange( cursor2, False )
                        controller.insertTransferable( txt )
                        i += 1
            return i
        except:
            return i
    def duplicateRow(self, phrase, regex=False):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        #when the cursor in is a table, the elements in the enumeration are tables, and not cells like I was expecting
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            e = x.nextElement()
            cellNames = e.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = e.getCellByName( cellName )
                text = cell.Text.getString()
                if text == phrase:
                    #we found the cell, now get the position of the cell ie b2
                    rowpos, colpos = self._convertCellNameToCellPositions(cellName)
                    rows = e.getRows()
                    #insert new row
                    rows.insertByIndex( rowpos+1, 1 )
                    self._copyRowCells( e, cellName )
                    return
    def _copyRowCells( self, table, cellName ):
        #start by getting the current row number
        matches = re.match( "(\w)+(\d)+", cellName )
        row = matches.group( 2 )
        #initialize values
        cols = []
        cellCursor = None 
        #loop through all cells with this row number in the cell name
        for x in table.getCellNames():
            matches = re.match( "(\w)+(\d)+", x )
            if matches.group( 2 ) == row:
                if cellCursor is None:
                    #on first loop we need to get the cursor
                    cellCursor = table.createCursorByCellName( x )
                else:
                    #on every other loop we just need to move the cursor 1 position to the right
                    cellCursor.goRight( 1, False )
                #get text cursor for the cell and copy content
                currentCell = table.getCellByName( cellCursor.getRangeName() )
                cols.append( matches.group( 2 ) )
                #get text cursor for the new cell and paste content
                nextRow = int( matches.group(2) )+1
                cellDown = table.getCellByName( matches.group(1)+"%s" % nextRow )
                #this line is wicked important, without it the formatting DOES NOT work. 
                #as it turns out, the string cannot be blank.
                cellDown.setString( ' ' )
                #paste contents from source to targets
                controller = self.oodocument.getCurrentController()
                viewCursor = controller.getViewCursor()
                viewCursor.gotoRange( currentCell.Text, False )
                txt = controller.getTransferable()
                viewCursor.gotoRange( cellDown.Text, False )
                controller.insertTransferable( txt )
        return 
    def duplicateColumn( self, phrase, count=1, regex=False ):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        #when the cursor in is a table, the elements in the enumeration are tables, and not cells like I was expecting
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            e = x.nextElement()
            cellNames = e.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = e.getCellByName( cellName )
                text = cell.Text.getString()
                if text == phrase:
                    rowpos, colpos = self._convertCellNameToCellPositions(cellName)
                    cols = e.getColumns()
                    cols.insertByIndex( colpos+1, 1 )
                    self._copyColumnCells( e, cellName )
                    return
    def _copyColumnCells(self, table, cellName):
        #start by getting the current row number
        matches = re.match( "(\w)+(\d)+", cellName )
        col = matches.group( 1 )
        #initialize values
        cols = []
        cellCursor = None
        pos = 0 
        #loop through all cells with this row number in the cell name
        for x in table.getCellNames():
            matches = re.match( "(\w)+(\d)+", x )
            row = "%s" % B26.fromBase26( matches.group( 1 ) )
            col = "%s" % col
            if matches.group( 1 ) == col:
                currentcolumn = B26.fromBase26( matches.group(1) )
                x = matches.group( 1 ) + "%s" % row
                if cellCursor is None:
                    #on first loop we need to get the cursor
                    cellCursor = table.createCursorByCellName( x )
                else:
                    #on every other loop we just need to move the cursor 1 position down
                    cellCursor.goDown( 1, False )
                matches = re.match( "(\w)+(\d)+", cellCursor.getRangeName() )
                #get the source cell and copy content
                sourceCell = table.getCellByPosition( currentcolumn-1, int(matches.group( 2 ))-1 )
                #get target cell
                targetCell = table.getCellByPosition( currentcolumn, int(matches.group( 2 ))-1 )
                targetCell.setString( ' ' )
                #paste contents from source to targets
                '''a HUGE thanks goes to Alessandro Dentella for this solution which allows me to get 
                rid of the dependency on AutoText for this, which is awesome because AutoText requires 
                root and is very slow.  AutoText also seemed to give me a bunch of segmentation faults.
                http://stackoverflow.com/questions/4541081/openoffice-duplicating-rows-of-a-table-in-writer/4596191#4596191'''
                controller = self.oodocument.getCurrentController()
                viewCursor = controller.getViewCursor()
                viewCursor.gotoRange( sourceCell.Text, False )
                txt = controller.getTransferable()
                viewCursor.gotoRange( targetCell.Text, False )
                controller.insertTransferable( txt )
        return
#========static methods============================================================================
    @staticmethod
    def _convertCellNameToCellPositions( cellName ):
        matches = re.match( "(\w)+(\d)+", cellName )
        col = matches.group(1)
        row = int( matches.group(2) ) - 1
        col = ord( col ) - 65
        return ( row, col )
    @staticmethod
    def _base26Decode( num ):
        return int( num, 26 )
if __name__ == "__main__":
    from sys import argv, exit
    if len( argv ) == 3:
        WriterDocument.convert( argv[1], argv[2] )
class PermissionError( Exception ):
    value = None
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr( self.value )