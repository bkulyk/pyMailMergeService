#!/usr/bin/env python
import uno
import os
from sys import path
from com.sun.star.beans import PropertyValue
from com.sun.star.io import IOException
class OpenOfficeConnection:
    """Object to hold the connection to openoffice, so there is only ever one connection"""
    desktop = None
    host = "localhost"
    port = 8100
    @staticmethod
    def getConnection( **kwargs ):
        """Create the connection object to open office, only ever create one."""
        port = kwargs.get( 'port', OpenOfficeConnection.port )
        host = kwargs.get( 'host', OpenOfficeConnection.host )
        if OpenOfficeConnection.desktop is None:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
            context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
            OpenOfficeConnection.desktop = context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', context )
        return OpenOfficeConnection.desktop
class OpenOfficeDocument:
    """Represent an open office docuemnt, with some really dumbed down method names. 
    Exaple: saveAs instead of storetourl (or whatever)"""
    openoffice = None
    oodocument = None
    documentTypes = [ 'com.sun.star.text.TextDocument', 'com.sun.star.sheet.SpreadsheetDocument', 'com.sun.star.drawing.DrawingDocument', 'com.sun.star.presentation.PresentationDocument' ]
    #there must be a programatic way to get this stuff instead of hardcoding, I managed to get the filter list from filterFactory... but can't figure out where to get the extensions.
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
                                'com.sun.star.text.TextDocument':'HTML (StarWriter)',
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
    @staticmethod
    def _makeProperty( key, value ):
        """Create an property for use with the OpenOffice API"""
        property = PropertyValue()
        property.Name = key
        property.Value = value
        return property
    def _getExportFilter( self, filename ):
        """Automatically determine to output filter depending on the file extension"""
        #self._makeProperty( 'FilterName', 'writer_pdf_Export' )
        ext = OpenOfficeDocument._getFileExtension( filename )
        for x in OpenOfficeDocument.documentTypes:
            if self.oodocument.supportsService( x ):
                if x in OpenOfficeDocument.exportFilters[ext].keys():
                    return OpenOfficeDocument._makeProperty( 'FilterName', OpenOfficeDocument.exportFilters[ext][x] )
        return None
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
if __name__ == "__main__":
    from sys import argv, exit
    if len( argv ) == 3:
        OpenOfficeDocument.convert( argv[1], argv[2] )