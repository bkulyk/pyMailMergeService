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
        sheet = self._getSheets()[ 0 ]
        range = sheet.getCellRangeByPosition( address.StartRow, address.StartColumn, address.EndRow, address.EndColumn )
        #get the rows
        rows = sheet.getRows()
        #add new row after the ending of the search term
        rows.insertByIndex( address.EndRow+1, 1 )
        #now copy contents of original row to new row
        #@todo
    def searchAndReplaceFirst( self, phrase, replacement, regex=False ):
        pass
    def refresh(self):
        pass