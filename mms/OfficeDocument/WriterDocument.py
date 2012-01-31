#!/usr/bin/env python
import uno, os, re, traceback, sys
from mms.OfficeDocument import OfficeDocument

from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
from com.sun.star.style.BreakType import PAGE_BEFORE, PAGE_AFTER, NONE
from com.sun.star.style.NumberingType import ARABIC
from com.sun.star.text.PageNumberType import CURRENT
from com.sun.star.style.ParagraphAdjust import RIGHT
from com.sun.star.chart.ChartDataRowSource import COLUMNS

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
                cursor.BreakType = PAGE_AFTER
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
        if cursor is not None and replacement is not None:
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
            if result1 is not None and result2 is not None:
                #create new cursor at start of first phrase, expand to second, and return the cursor range
                cursor = self.oodocument.Text.createTextCursorByRange( result.getStart() )
                cursor.gotoRange( result2.getEnd(), True )
                return cursor
        except:
            pass
        return None
    def searchAndRemoveSection( self, startPhrase, endPhrase, regex=False, debug=False ):
        cursor = self._getCursorForStartAndEndPhrases(startPhrase, endPhrase, regex)
        if cursor is None:
            #it seems on occasion the _getCursorForStartAndEndPhrases way or doing this does not work
            c1 = self._getCursorForStartPhrase( startPhrase )
            c2 = self._getCursorForStartPhrase( endPhrase )
            if c1 is not None and c2 is not None:
                c1.gotoRange( c2.getEnd(), True )
                c1.setString( '' )
                return 1
            return 0
        else:
            self.oodocument.Text.insertString( cursor, '', True )
            return 1
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
        
    def deleteRow( self, phrase ):
        cursor = self._getCursorForStartPhrase( phrase )
        if cursor is None:
            return 0
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            table = x.nextElement()
            cellNames = table.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = table.getCellByName( cellName )
                if cell is not None:
                    text = cell.Text.getString()
                    if phrase in text: #text == phrase:
                        sourceRow = self._convertCellNameToCellPositions( cellName )[0]
                        table.Rows.removeByIndex( sourceRow, 1 )
                        return 1
        return 0
    
    def deleteColumn(  self, phrase ):
        cursor = self._getCursorForStartPhrase( phrase )
        if cursor is None:
            return 0
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            table = x.nextElement()
            cellNames = table.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = table.getCellByName( cellName )
                if cell is not None:
                    text = cell.Text.getString()
                    if phrase in text: #text == phrase:
                        sourceCol = self._convertCellNameToCellPositions( cellName )[1]
                        table.Columns.removeByIndex( sourceCol, 1 )
                        return 1
        return 0
                    
    def duplicateRow( self, phrase, count=1, regex=False ):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        if cursor is None:
            return 0
        #when the cursor in is a table, the elements in the enumeration are tables, and not cells like I was expecting
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            table = x.nextElement()
            cellNames = table.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = table.getCellByName( cellName )
                text = cell.Text.getString()
                if text == phrase:
                    #we found the source row, copy the cell contents, and insert new row and paste contents
                    ##copy contents
                    sourceRow = self._convertCellNameToCellPositions( cellName )[0]
                    row = table.Rows.getByIndex( sourceRow )
                    sourceCols = []
                    for colIndex in xrange( table.Columns.getCount() ):
                        cell = table.getCellByPosition( colIndex, sourceRow )
                        #get cell contents
                        controller = self.oodocument.getCurrentController()
                        viewCursor = controller.getViewCursor()
                        try:
                            viewCursor.gotoRange( cell.Text, False )
                            txt = controller.getTransferable()
                            sourceCols.append( txt )
                        except:
                            traceback.print_exc( file=sys.stdout )
                            print 'cant copy cells. phrase: %s text: %s ---- %s-%s' % ( phrase, text, colIndex, sourceRow )
                            
                    #now insert new row and paste content into said row.
                    rows = table.getRows()
                    for i in xrange( count ):
                        rows.insertByIndex( sourceRow+1, 1 )
                        for colIndex in xrange( table.Columns.getCount() ):
                            try:
                                cell = table.getCellByPosition( colIndex, sourceRow+1 )
                                viewCursor.gotoRange( cell.Text, False )
                                controller.insertTransferable( sourceCols[ colIndex ] )
                            except:
                                print 'cant paste cells. phrase: %s text: %s ---- %s-%s' % ( phrase, text, colIndex, sourceRow )
    
    def duplicateColumn( self, phrase, count=1, regex=False ):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        #when the cursor in is a table, the elements in the enumeration are tables, and not cells like I was expecting
        x = cursor.createEnumeration()
        if x.hasMoreElements():
            table = x.nextElement()
            cellNames = table.getCellNames()
            #need to find the cell with the search phrase that was provided
            for cellName in cellNames:
                cell = table.getCellByName( cellName )
                text = cell.Text.getString()
                if text == phrase:
                    
                    #we found the source cell, copy the cell contents, and insert new column and paste contents
                    sourceCol = self._convertCellNameToCellPositions( cellName )[1]
                    sourceRows = []
                    for rowIndex in xrange( table.Rows.getCount() ):
                        cell = table.getCellByPosition( sourceCol, rowIndex )
                        
                        #get cell contents
                        controller = self.oodocument.getCurrentController()
                        viewCursor = controller.getViewCursor()
                        viewCursor.gotoRange( cell.Text, False )
                        txt = controller.getTransferable()
                        sourceRows.append( txt )
                        
                    #now insert new column and paste content into said column.
                    for i in xrange( count ):
                        cols = table.getColumns()
                        cols.insertByIndex( sourceCol+1+i, 1 )
                        cell = table.getCellByPosition( sourceCol+1+i, 0 )
                        cell.Text.setString( i )
                        for rowIndex in xrange( len( sourceRows ) ):
                            cell = table.getCellByPosition( sourceCol+1+i, rowIndex )
                            viewCursor.gotoRange( cell.Text, False )
                            controller.insertTransferable( sourceRows[ rowIndex ] )
                    return
                    
    def getTextTableStrings( self, table ):
        #extract data
        tableData = []
        for row in xrange( table.getRows().getCount() ):
            rowData = []
            for col in xrange( table.getColumns().getCount() ):
                cell = table.getCellByPosition( col, row )
                rowData.append( cell.Text.getString() )
            tableData.append( rowData )
        return tableData
    
    def updateCharts( self ):
        #get all of the tables in the document, we'll need them later
        tables = self.oodocument.getTextTables()
        
        #loop through each of the draw pages in this document
        count = 0
        drawPage = self.oodocument.getDrawPage()
        e = drawPage.createEnumeration()
        while e.hasMoreElements():
            x = e.nextElement()
            count += 1
            
            try:
                #get the chart object
                embedded = x.getEmbeddedObject()
                skip = False
            except:
                skip = True
            
            if skip is False:
                #get a tuple of all the ranges used in the chart
                used = embedded.getUsedRangeRepresentations()
                #eg, (u"Table1.A2:A4", u"Table1.B1:B1", u"Table1.B2:B4", u"Table1.C1:C1", u"Table1.C2:C4" )
                
                #need to pass these properties to the new data source we will create
                props = 'DataSourceLabelsInFirstRow', 'DataSourceLabelsInFirstColumn'
                DataSourceLabelsInFirstRow, DataSourceLabelsInFirstColumn = embedded.getPropertyValues( props )
    
                #get the table the range is based off of
                tablename = used[0].split('.')[0]
                try:
                    table = tables.getByName( tablename )
                except:
                    table = None
                
                if table is not None:
                    #determine the cell name that we are going to start out new range with
                    cellName = used[0].split('.')[1].split(":")[0]
                    row, col = self._convertCellNameToCellPositions( cellName )
                    if DataSourceLabelsInFirstRow and row > 1:
                        row = int(row)-1
                    startCellName =  "%s%s" % ( re.sub( '\d+', '', cellName ), row )
                    
                    #determine the cell name that we are going to end our new range with
                    lastcolumn = used[ len(used)-1 ].split(":")[1]
                    columnName = re.sub( '\d+', '', lastcolumn )
                    
                    #build the representation string ie.  "Table1.A1:C4"
                    representation = "%s.%s:%s%s" % (tablename, startCellName, columnName, table.Rows.getCount() )
                    
                    #create a new re.sub( '\d+', '', cellName ) out of the representation we built
                    dp = embedded.getDataProvider()
                    ds = WriterDocument.createDataSource( dp, representation, COLUMNS, DataSourceLabelsInFirstRow, DataSourceLabelsInFirstColumn )
                    
                    #get the diagram object and tell it to use our new data source
                    diagram = embedded.getFirstDiagram()
                    
                    try:
                        #this is a much better way of updating the chart, however I have a system
                        #  where the setDiagramData method does not exist, and throws an exception
                        #  never fear.... I have a backup plan.
                        prop = OfficeDocument._makeProperty( "HasCategories", True ),
                        diagram.setDiagramData( ds, prop )
                    except:
                        #here is the backup plan, it's a lot more convoluted
                        sequences = ds.getDataSequences()
        
                        new_data = []
                        for x in sequences:
                            values = []
                            for x in x.getValues().getData():
                                try:
                                    values.append( float( x.replace('$','').replace( ',', ''  ).replace( ' ', ''  ) ) )
                                except:
                                    pass
                            if len( values ) > 0:
                                new_data.append( tuple( values ) )
        
                        #need to invert the matrix 
                        new_data = WriterDocument.invertTuple( new_data )
        
                        #set the data
                        embedded.getData().setData( new_data )
                        
                        #new get the data for the updated legends
                        pos_row, pos_col = WriterDocument._convertCellNameToCellPositions( used[0].split('.')[1] )
                        new_legends = []
                        for x in xrange( pos_row-1, table.getRows().getCount()-1 ):
                            new_legends.append( table.getCellByPosition( pos_col, x+pos_row ).Text.getString() )
                        
                        #set the updated legend data
                        embedded.getData().setRowDescriptions( tuple(new_legends) )
            
    def refresh(self):
        OfficeDocument.refresh( self )
        try:
            self.updateCharts()
        except:
            traceback.print_exc( file=sys.stdout )
            pass
        
#========static methods============================================================================
    @staticmethod 
    def toNumber( input_string ):
        return float( re.sub( r"[^\d+\.]", '', input_string ) )
    
    @staticmethod
    def invertTuple( data ):
        new = []
        for x in xrange( len( data[0] ) ):
            vals = []
            for y in xrange( len( data ) ):
                vals.append( data[y][x] )
            new.append( tuple( vals ) )
        
        return tuple( new )
    
    @staticmethod
    def _convertCellNameToCellPositions( cellName ):
        matches = re.match( "(\w)+(\d)+", cellName )
        col = matches.group(1)
        row = int( matches.group(2) ) - 1
        col = ord( col ) - 65
        return ( row, col )
    
    @staticmethod
    def createDataSource( DataProvider, CellRangeRepresentation, DataRowSource, FirstCellAsLabel=True, HasCategories=True ):
        props = []
        props.append( OfficeDocument._makeProperty( 'CellRangeRepresentation', CellRangeRepresentation ) )
        props.append( OfficeDocument._makeProperty( 'DataRowSource', DataRowSource ) )
        props.append( OfficeDocument._makeProperty( 'FirstCellAsLabel', FirstCellAsLabel ) )
        props.append( OfficeDocument._makeProperty( 'HasCategories', HasCategories ) )
        
        props = tuple( props )
        
        ds = DataProvider.createDataSource( props )
        return ds
    
if __name__ == "__main__":
    from sys import argv, exit
    if len( argv ) == 3:
        WriterDocument.convert( argv[1], argv[2] )
        