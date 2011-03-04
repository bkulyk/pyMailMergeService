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
    def _debugMethod( self, unoobj, methodName ):
        from com.sun.star.beans.MethodConcept import ALL as ALLMETHS
        from com.sun.star.beans.PropertyConcept import ALL as ALLPROPS
        ctx = OpenOfficeConnection.context
        introspection = ctx.ServiceManager.createInstanceWithContext( "com.sun.star.beans.Introspection", ctx)
        access = introspection.inspect(unoobj)
        method = access.getMethod( methodName, ALLMETHS )
        for x in method.getParameterInfos():
            print x
        print ""
    def _debug( self, unoobj, doPrint=False ):
        """
        Print(or return) all of the method and property names for an uno object6
        Thanks to Carsten Haese for his insanely useful example fount at:
        http://bytes.com/topic/python/answers/641662-how-do-i-get-type-methods#post2545353
        """
        from com.sun.star.beans.MethodConcept import ALL as ALLMETHS
        from com.sun.star.beans.PropertyConcept import ALL as ALLPROPS
        ctx = OpenOfficeConnection.context
        introspection = ctx.ServiceManager.createInstanceWithContext( "com.sun.star.beans.Introspection", ctx)
        access = introspection.inspect(unoobj)
        meths = access.getMethods(ALLMETHS)
        props = access.getProperties(ALLPROPS)
        if doPrint:
            print "Object Methods:"
            for x in meths:
                print "---- %s" % x.getName()
            print "Object Properties:"
            for x in props:
                print "---- %s" % x.Name
        return [ x.getName() for x in meths ] + [ x.Name for x in props ]
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
    def duplicateRow(self, phrase, regex=False):
#        print "phrase: %s" % phrase
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
                textCursor = currentCell.createTextCursor()
                textCursor.gotoStart( False )
                textCursor.gotoEnd( True )
                at = AutoText( textCursor )
                cols.append( matches.group( 2 ) )
#                print "copy text from: %s" %  cellCursor.getRangeName()
                #get text cursor for the new cell and paste content
                nextRow = int( matches.group(2) )+1
                cellDown = table.getCellByName( matches.group(1)+"%s" % nextRow )
                cellDownTextCursor = cellDown.createTextCursor()
                cellDownTextCursor.gotoStart( False )
                cellDownTextCursor.gotoEnd( True )
                cellDownCellCursor = table.createCursorByCellName( "%s%s" % ( matches.group( 1 ), nextRow  ) )
                at.insert( cellDownTextCursor )
                at.delete()
#                print "copy text to: %s%s" % ( matches.group(1) , nextRow )
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
#                print "Copy cell from: col %s, row %s" % ( currentcolumn-1, int(matches.group( 2 ))-1 )
                sourceCell = table.getCellByPosition( currentcolumn-1, int(matches.group( 2 ))-1 )
                sourceCursor = sourceCell.createTextCursor()
                sourceCursor.gotoEnd( True )
                at = AutoText( sourceCursor )
                #get target cell and paste contents
                targetCell = table.getCellByPosition( currentcolumn, int(matches.group( 2 ))-1 )
#                print "try to Copy cell to: col %s, row %s" % ( currentcolumn, int(matches.group( 2 ))-1 )
                targetCursor = targetCell.createTextCursor()
                at.insert( targetCursor )
                #delete the autotext entry now that we don't need it any longer
                at.delete()
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