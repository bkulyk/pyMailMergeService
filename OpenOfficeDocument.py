#!/usr/bin/env python
import uno
import os
from sys import path
from com.sun.star.beans import PropertyValue
from com.sun.star.io import IOException
import uuid
import re
class OpenOfficeConnection:
    """Object to hold the connection to openoffice, so there is only ever one connection"""
    desktop = None
    host = "localhost"
    port = 8100
    context = ''
    @staticmethod
    def getConnection( **kwargs ):
        """Create the connection object to open office, only ever create one."""
        port = kwargs.get( 'port', OpenOfficeConnection.port )
        host = kwargs.get( 'host', OpenOfficeConnection.host )
        if OpenOfficeConnection.desktop is None:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
            OpenOfficeConnection.context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
            OpenOfficeConnection.desktop = OpenOfficeConnection.context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', OpenOfficeConnection.context )
        return OpenOfficeConnection.desktop
class AutoText:
    autoTextName = None
    autoTextGroup = None
    autoTextContainer = None
    autoTextEntry = None
    def __init__( self, cursor ):
        try:
            #create a new auto text container
            self.autoTextContainer = OpenOfficeConnection.context.ServiceManager.createInstanceWithContext( 'com.sun.star.text.AutoTextContainer', OpenOfficeConnection.context )
            #create a unique name for the container
            self.autoTextName = "%s" % uuid.uuid1()
            self.autoTextName = self.autoTextName.replace( '-', '_' ) #will only accept a-z, A-Z spaces and underscores
            if self.autoTextContainer.hasByName( self.autoTextName ):
                self.autoTextContainer.removeByName( self.autoTextName )
            self.autoTextGroup = self.autoTextContainer.insertNewByName( self.autoTextName )
            self.autoTextEntry = self.autoTextGroup.insertNewByName( 'MAE', 'My AutoText Entry', cursor )
        except:
            raise PermissionError( "You likely do not have premission needed to create an autotext entry. (needed for copy/paste functions)  Try running soffice as root." )
    def insert( self, cursor ):
        self.autoTextEntry.applyTo( cursor )
    def getTextContent(self):
        return self.autoTextEntry
    def delete( self ):
        self.autoTextContainer.removeByName( self.autoTextName )
        pass
class OpenOfficeDocument:
    """Represent an open office document, with some really dumbed down method names. 
    Example: saveAs instead of storetourl (or whatever)"""
    openoffice = None
    oodocument = None
    documentTypes = [ 'com.sun.star.text.TextDocument', 'com.sun.star.sheet.SpreadsheetDocument', 'com.sun.star.drawing.DrawingDocument', 'com.sun.star.presentation.PresentationDocument' ]
    #there must be a programmatic way to get this stuff instead of hard-coding, I managed to get the filter list from filterFactory... but can't figure out where to get the extensions.
    #list pulled from: http://wiki.services.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
    #I'm probably missing many types.
    exportFilters = { 
                     #general
                         'pdf':{
                                'com.sun.star.text.TextDocument':'writer_pdf_Export',
                                'com.sun.star.sheet.SpreadsheetDocument':'calc_pdf_Export',
                                'com.sun.star.drawing.DrawingDocument':'draw_pdf_Export',
                                'com.sun.star.presentation.PresentationDocument':'impress_pdf_Export',
                                'com.sun.star.text.WebDocument':'writer_web_pdf_Export',
                                'com.sun.star.formula.FormulaProperties':'math_pdf_Export',
                                'com.sun.star.text.GlobalDocument':'writer_globaldocument_pdf_Export'
                                },
                        'html':{
                                'com.sun.star.text.TextDocument':'writer8',
                                #'com.sun.star.text.TextDocument':'HTML (StarWriter)',
                                'com.sun.star.text.WebDocument':'HTML',
                                'com.sun.star.text.GlobalDocument':'writerglobal8_HTML',
                                'com.sun.star.sheet.SpreadsheetDocument':'HTML (StarCalc)',
                                'com.sun.star.drawing.DrawingDocument':'draw_html_Export',
                                'com.sun.star.presentation.PresentationDocument':'impress_html_Export'
                                },
                        'xhtml':{
                                 'com.sun.star.text.TextDocument':'XHTML Writer File',
                                 'com.sun.star.sheet.SpreadsheetDocument':'XHTML Calc File',
                                 'com.sun.star.drawing.DrawingDocument':'XHTML Draw File',
                                 'com.sun.star.presentation.PresentationDocument':'XHTML Impress File'
                            },
                    #writer specific
                        'doc':{
                                "com.sun.star.text.TextDocument":"MS Word 97"
                                },
                        'docx':{
                                "com.sun.star.text.TextDocument":"MS Word 2007 XML"
                                },
                        'odt':{
                               'com.sun.star.text.TextDocument':'writer8',
                               'com.sun.star.text.WebDocument':'writerweb8_writer',
                               'com.sun.star.text.GlobalDocument':'writer_globaldocument_StarOffice_XML_Writer_GlobalDocument'
                               },
                        'wiki':{
                                'com.sun.star.text.TextDocument':'MediaWiki',
                                'com.sun.star.text.WebDocument':'MediaWiki_Web'
                                },
                        'txt':{
                               'com.sun.star.text.TextDocument':'Text',
                               'com.sun.star.text.WebDocument':'Text (StarWriter/Web)'
                               },
                        'rtf':{
                               'com.sun.star.text.TextDocument':'Rich Text Format',
                               'com.sun.star.sheet.SpreadsheetDocument':'Rich Text Format (StarCalc)'
                               },
                        'ods':{
                               'com.sun.star.sheet.SpreadsheetDocument':'calc8'
                               },
                    #calc specific
                        'xls':{
                               'com.sun.star.sheet.SpreadsheetDocument':'MS Excel 97'
                               },
                        'xlsx':{
                                'com.sun.star.sheet.SpreadsheetDocument':'Calc MS Excel 2007 XML'
                                },
                    #impress specific
                        'pptx':{
                                'com.sun.star.presentation.PresentationDocument':'Impress MS PowerPoint 2007 XML'
                                },
                        'ppt':{
                               'com.sun.star.presentation.PresentationDocument':'MS PowerPoint 97'
                               },
                        'odp':{
                               'com.sun.star.presentation.PresentationDocument':'impress8'
                               },
                        'svg':{
                               'com.sun.star.drawing.DrawingDocument':'draw_svg_Export',
                               'com.sun.star.presentation.PresentationDocument':'impress_svg_Export'
                               }
                    }
    def __init__(self):
        """Connect to openoffice"""
        self.openoffice = OpenOfficeConnection.getConnection()
    def open( self, filename ):
        """Open an OpenOffice document"""
        #http://www.oooforum.org/forum/viewtopic.phtml?t=35344
        properties = []
        properties.append( OpenOfficeDocument._makeProperty( 'Hidden', True ) ) 
        properties = tuple( properties )
        self.oodocument = self.openoffice.loadComponentFromURL( uno.systemPathToFileUrl( os.path.abspath( filename ) ), "_blank", 0, properties )
    def refresh( self, refreshIndexes=True ):
        """Refresh the OpenOffice docuemnt and (optionally) it's indexes"""
        try:
            self.oodocument.refresh()
        except:
            pass
        if refreshIndexes:
            #I needed the table of contents to automatically update in case page contents had changed
            try:
                #get all document indexes, eg. toc, or index
                oIndexes = document.getDocumentIndexes()
                for x in range( 0, oIndexes.getCount() ):
                    oIndexes.getByIndex( x ).update()
            except:
                pass
    def saveAs( self, filename ):
        """Save the open office document to a new file, and possibly filetype. 
        The type of document is parsed out of the file extension of the filename given."""
        filename = uno.systemPathToFileUrl( os.path.abspath( filename ) )
        #filterlist: http://wiki.services.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
        exportFilter = self._getExportFilter( filename )
        props = exportFilter, 
        #storeToURL: #http://codesnippets.services.openoffice.org/Office/Office.ConvertDocuments.snip
        self.oodocument.storeToURL( filename, props )
    def close( self ):
        """Close the OpenOffice document"""
        self.oodocument.close( 1 )
    def _getExportFilter( self, filename ):
        """Automatically determine to output filter depending on the file extension"""
        #self._makeProperty( 'FilterName', 'writer_pdf_Export' )
        ext = OpenOfficeDocument._getFileExtension( filename )
        for x in OpenOfficeDocument.documentTypes:
            if self.oodocument.supportsService( x ):
                if x in OpenOfficeDocument.exportFilters[ext].keys():
                    return OpenOfficeDocument._makeProperty( 'FilterName', OpenOfficeDocument.exportFilters[ext][x] )
        return None
    def searchAndReplaceWithDocument( self, phrase, documentPath, regex=False ):
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_1_1_Editing_Text
    	#cursor = self.oodocument.Text.createTextCursor()
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_3_3_Search_and_Replace
    	search = self.oodocument.createSearchDescriptor()
    	search.setSearchString( phrase )
    	search.SearchRegularExpression = regex
    	result = self.oodocument.findFirst( search )
    	path = uno.systemPathToFileURL( os.path.abspath( path ) )
        #http://api.openoffice.org/docs/DevelopersGuide/Text/Text.xhtml#1_3_1_5_Inserting_Text_Files
    	return result.insertDocumentFromURL( path, tuple() )
    def searchAndReplace( self, phrase, replacement, regex=False ):
        #cursor = self.oodocument.Text.createTextCursor()
        replace = self.oodocument.createReplaceDescriptor()
        replace.setSearchString( phrase )
        replace.setReplaceString( replacement )
        replace.SearchRegularExpression = regex
        return self.oodocument.replaceAll( replace )
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
        #create new cursor at start of first phrase, expand to second, and return the cursor range
        cursor = self.oodocument.Text.createTextCursorByRange( result.getStart() )
        cursor.gotoRange( result2.getEnd(), True )
        return cursor
    def searchAndRemoveSection( self, startPhrase, endPhrase, regex=False ):
        cursor = self._getCursorForStartAndEndPhrases(startPhrase, endPhrase, regex)
        self.oodocument.Text.insertString( cursor, '', True )
    def searchAndDuplicate( self, startPhrase, endPhrase, count, regex=False ):
        cursor = self._getCursorForStartAndEndPhrases( startPhrase, endPhrase, regex )
        cursor2 = self.oodocument.Text.createTextCursorByRange( cursor.getEnd() )
        at = AutoText( cursor )
        for x in xrange( count ):
            at.insert( cursor2 )
        at.delete()
    def duplicateRow(self, phrase, count=1, regex=False):
        cursor = self._getCursorForStartPhrase( phrase, regex )
        at = AutoText( cursor )
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
                    rows.insertByIndex( rowpos, 1 )
                    #highlight/select the entire content of the original row, so that it's content can be copied and pasted to the new row
                    tableCursor = e.createCursorByCellName( rowpos )
#                    tableCursor.gotoEnd( True )
#                    print cell.Text.getString()
#                    print cell
#                    print "======"
                    self._copyCells( e, rowpos, rowpos+1 )
                    
#                    cell2 = e.getCellByName( "B2" )
#                    print cell2
#                    cellCursor2 = cell2.createTextCursor()
#                    
#                    cellCursor = cell.createTextCursor()
#                    cellCursor.gotoEnd( True )
#                    at.insert( cellCursor2 )
#                    at.delete()
                    return
                    
#                    cursor = self.oodocument.Text.createTextCursor()
#                    self.oodocument.Text.insertString( cursor, 'here', 0 )
    def _copyCells( self, table, fromRowIndex, toRowIndex ):
        return
#        print table
        #find out how many columns are in the fromRow using the column names. ie if it's position 2, find out how many cell's names start with B  
        print table.getCellNames()
        
        TableRows = table.getRows()
        fromTableRow = TableRows.getByIndex( fromRowIndex )
        print TableRows.getCount()
        
        TableColumns = table.getColumns()
        print TableColumns.getCount()
        
        cell = table.getCellByPosition( fromRowIndex, 0 )
#        print cell
#========static methods============================================================================
    @staticmethod
    def _convertCellNameToCellPositions( cellName ):
        matches = re.match( "(\w)+(\d)+", cellName )
        row = matches.group(1)
        col = int( matches.group(2) ) - 1
        #convert the letter to a number
        #@todo the following will need to be adjusted for cell names like aa4 
        
#        
#        newString = ""
#        for x in row:
#            tmp = ord( row ) - 64
#            tmp = 
            
        row = ord( row ) - 65
        return ( row, col )
    @staticmethod
    def _base26Decode( num ):
        return int( num, 26 )
    @staticmethod
    def _getFileExtension( filepath ):
        """Get the file extension for the given path"""
        file = os.path.splitext(filepath.lower())
        if len( file ):
            return file[1].replace( '.', '' )
        else:
            return filepath
    @staticmethod
    def convert( inputFile, outputFile ):
        """Convert the given input file to whatever type of file the outputFile is."""
        c = OpenOfficeDocument()
        c.open( inputFile )
        c.refresh()
        c.saveAs( outputFile )
        c.close()
    @staticmethod
    def _makeProperty( key, value ):
        """Create an property for use with the OpenOffice API"""
        property = PropertyValue()
        property.Name = key
        property.Value = value
        return property
if __name__ == "__main__":
    from sys import argv, exit
    if len( argv ) == 3:
        OpenOfficeDocument.convert( argv[1], argv[2] )
class PermissionError( Exception ):
    value = None
    def __init__(self, value):
        self.value = value
    def __str__(self):
        return repr( self.value )
class B26:
    @staticmethod
    def fromBase26( value ):
        total = 0
        pos = 0
        for digit in value[::-1]:
            dec = int( digit, 36 ) - 9
            if pos == 0:
                add = dec
            else:
                x = dec * pos
                add = x * 26
            total = total + add
            pos = pos + 1
        return total
    @staticmethod
    def toBase26( number ):
        """
        Convert positive integer to a base36 string.
        I gave up trying to write this myself, and grabbed it from Wikipedia's base 36 convert example
        http://en.wikipedia.org/wiki/Base_36#Python_Conversion_Code 
        """
        alphabet='0123456789ABCDEFGHIJKLMNOP'
        if not isinstance(number, (int, long)):
            raise TypeError('number must be an integer')
        # Special case for zero
        if number == 0:
            return '0'
        base36 = ''
        sign = ''
        if number < 0:
            sign = '-'
            number = - number
        while number != 0:
            number, i = divmod(number, len(alphabet))
            base36 = alphabet[i] + base36
        #real base 26 is 0 to P, we need it to be A-Z, so adjust each char accordingly
        alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        c = sign + base36
        string = ''
        count = 0
        for d in c[::-1]:
            if count == 0:
                y = alpha[ int( d, 26 )-1 ]
            else:
                y = alpha[ int( d, 26 )-1 ]
            string = "%s%s" % ( y, string )
            count = count + 1
        return string.replace( "AZ", "Z")