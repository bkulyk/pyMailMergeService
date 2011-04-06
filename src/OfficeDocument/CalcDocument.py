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
    def easydebug( self, obj ):
        meths, props = self._debug( obj )
        meths.sort()
        for x in meths:
            print '----- ' + x
        print ''
    def searchAndReplaceFirst( self, phrase, replacement, regex=False ):
        pass
    def refresh(self):
        pass