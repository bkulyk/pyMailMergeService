import uno
import os
from sys import path
from com.sun.star.beans import PropertyValue
from com.sun.star.io import IOException
from com.sun.star.document.MacroExecMode import NEVER_EXECUTE #, FROM_LIST, ALWAYS_EXECUTE, USE_CONFIG, ALWAYS_EXECUTE_NO_WARN, USE_CONFIG_REJECT_CONFIRMATION, USE_CONFIG_APPROVE_CONFIRMATION, FROM_LIST_NO_WARN, FROM_LIST_AND_SIGNED_WARN, FROM_LIST_AND_SIGNED_NO_WARN
class OfficeConnection:
    """Object to hold the connection to open office/libre office, so there is only ever one connection"""
    desktop = None
    host = "localhost"
    port = 8100
    context = ''
    @staticmethod
    def getConnection( **kwargs ):
        """Create the connection object to open office, only ever create one."""
        port = kwargs.get( 'port', OfficeConnection.port )
        host = kwargs.get( 'host', OfficeConnection.host )
        if OfficeConnection.desktop is None:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
            OfficeConnection.context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
            OfficeConnection.desktop = OfficeConnection.context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', OfficeConnection.context )
        return OfficeConnection.desktop
class OfficeDocument:
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
    def __init__( self ):
        """Connect to openoffice"""
        self.openoffice = OfficeConnection.getConnection()
    def open( self, filename ):
        """Open an OpenOffice document"""
        #http://www.oooforum.org/forum/viewtopic.phtml?t=35344
        properties = []
        properties.append( OfficeDocument._makeProperty( 'Hidden', True ) )
        properties.append( OfficeDocument._makeProperty( 'MacroExecutionMode', NEVER_EXECUTE ) )
        properties.append( OfficeDocument._makeProperty( 'ReadOnly', False ) )
        properties = tuple( properties )
        self.oodocument = self.openoffice.loadComponentFromURL( uno.systemPathToFileUrl( os.path.abspath( filename ) ), "_blank", 0, properties )
    @staticmethod
    def createDocument( type ):
        if type == 'odt':
            from OfficeDocument.WriterDocument import WriterDocument
            return WriterDocument()
        elif type == 'ods':
            from OfficeDocument.CalcDocument import CalcDocument
            return CalcDocument()
    @staticmethod
    def _makeProperty( key, value ):
        """Create an property for use with the OpenOffice API"""
        property = PropertyValue()
        property.Name = key
        property.Value = value
        return property