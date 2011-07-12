from OfficeDocument import OfficeDocument
class CalcDocument( OfficeDocument ):
    __sheets = []
    def _getSheets(self):
        if len( self.__sheets ) == 0:
            sheets = self.oodocument.getSheets().createEnumeration()
            while sheets.hasMoreElements():
                sheet = sheets.nextElement()
                if sheet is not None:
                    self.__sheets.append( sheet )
                #find position of start phrase
        return self.__sheets
    def _getCellForStartPhrase( self, startPhrase, regex=False ):
        sheets = self._getSheets()
        for sheet in sheets:
            search = sheet.createSearchDescriptor()
            search.setSearchString( startPhrase )
            search.SearchRegularExpression = regex
            try:
                result = sheet.findFirst( search )
                if result is not None:
                    return result
            except:
                pass
        return None
    def duplicateRow( self, phrase, regex=False ):
        #get the cell address 
        cell = self._getCellForStartPhrase( phrase )        
        address = cell.getRangeAddress()
        #create a cell range
        sheet = self._getSheets()[ address.Sheet ]
        range = sheet.getCellRangeByPosition( address.StartColumn, address.StartRow,address.StartColumn, address.StartRow )# address.EndColumn, address.EndColumn )
        #get the rows
        rows = sheet.getRows()
        #add new row after the ending of the search term
        rows.insertByIndex( address.EndRow+1, 1 )
        #get max used area
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea( False )
        maxaddress = cursor.getRangeAddress()
        #get the entire row
        rangeToCopy = sheet.getCellRangeByPosition( address.StartColumn, address.StartRow, maxaddress.EndColumn, address.StartRow )
        #get entire next row
        rangeToPaste = sheet.getCellRangeByPosition( address.StartColumn, address.StartRow+1, maxaddress.EndColumn, address.StartRow+1 )
        #copy data from first row to second row
        rangeToPaste.setDataArray( rangeToCopy.getDataArray() )
    
    def searchAndReplaceFirst( self, phrase, replacement, regex=False ):
        cell = self._getCellForStartPhrase( phrase )
        '''if I don't convert numbers into numbers than formula don't work, and 
        Excel gives warnings about numbers being formatted as strings'''
        try:
            #with throw an exception for floating points and strings
            replacement = int(replacement)
        except:
            try:
                #will throw an exception for strings
                replacement = float(replacement)
            except:
                pass
        replacementArray = ( replacement, ),
        cell.setDataArray( replacementArray )
    
    def searchAndReplace( self, phrase, replacement, regex=False ):
        sheets = self._getSheets()
        count = 0
        for sheet in sheets:
            replace = sheet.createReplaceDescriptor()
            replace.setSearchString( phrase )
            replace.setReplaceString( replacement )
            replace.SearchRegularExpression = regex
            count += sheet.replaceAll( replace )
        return count
    
    def refresh(self):
        pass
    
    def getNamedRanges(self):
        return self.oodocument.NamedRanges.ElementNames
    
    def getNamedRangeData( self, rangeName ):
        namedRange = self.oodocument.NamedRanges.getByName( rangeName )
        sheetAndRangeDesc = namedRange.getContent() #ie. $Sheet1.$A$1:$C$4
        sheetName = sheetAndRangeDesc.split( '.' )[0][1:]
        sheet = self.oodocument.getSheets().getByName( sheetName )
        range = sheet.getCellRangeByName( sheetAndRangeDesc )
        data = list( range.getData() )
        #if there is many rows and only one column, make the data a list of values, instead of tuple of tuples
        i = 0
        for x in data:
            if len( x ) == 1:
                data[i] = x[0]
            i += 1
        return tuple( data )
    
    def getNamedRangeStrings( self, rangeName, limit=None ):
        namedRange = self.oodocument.NamedRanges.getByName( rangeName )
        sheetAndRangeDesc = namedRange.getContent() #ie. $Sheet1.$A$1:$C$4
        sheetName = sheetAndRangeDesc.split( '.' )[0][1:]
        sheet = self.oodocument.getSheets().getByName( sheetName )
        range = sheet.getCellRangeByName( sheetAndRangeDesc )
        
        #extract strings
        data = []
        i = 0
        quit = False
        for row in xrange( range.getRows().getCount() ):
            rowdata = []
            if quit is False:
                for col in xrange( range.getColumns().getCount() ):
                    if quit is False:
                        if limit is None or i < limit:
                            cell = range.getCellByPosition( col, row )
                            rowdata.append( cell.Text.getString() )
                            i += 1
                        else:
                            quit = True
                            break
                if len( rowdata ):
                    data.append( rowdata )
            
        #if there is only one row, only return a list instead of a list of lists
        if len( data ) == 1:
            return data[0]
        
        #if there is only one column, only return a list instead of a list of lists
        if len( rowdata ) == 1:
            i = 0
            for x in data:
                data[i] = data[i][0]
                i += 1
        return data
    
    def setNamedRangeStrings( self, rangeName, values ):
        namedRange = self.oodocument.NamedRanges.getByName( rangeName )
        sheetAndRangeDesc = namedRange.getContent() #ie. $Sheet1.$A$1:$C$4
        sheetName = sheetAndRangeDesc.split( '.' )[0][1:]
        sheet = self.oodocument.getSheets().getByName( sheetName )
        range = sheet.getCellRangeByName( sheetAndRangeDesc )
        
        rowcount = range.getRows().getCount()
        for row in xrange( rowcount ):
            colcount = range.getColumns().getCount()
            for col in xrange( colcount ):
                cell = range.getCellByPosition( col, row )
                
                if colcount > 1 and rowcount > 1:
                    value = values[ col ][ row ]
                elif colcount == 1 and rowcount > 1:
                    value = values[ row ]
                elif rowcount == 1 and colcount > 1:
                    value = values[ col ]
                    
                if CalcDocument.isNumber( value ):
                    cell.setValue( value )
                else:
                    cell.Text.setString( value )

    @staticmethod
    def isNumber( value ):
        return isinstance( value, ( int, float ) )