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
    filename = ''
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
    def createNew( self ):
        """Creates a new OpenOffice document"""
        return self.openUrl( "private:factory/swriter" )
    def open( self, filename ):
        """Open an OpenOffice document"""
        self.filename = os.path.abspath( filename )
        return self.openUrl( uno.systemPathToFileUrl( self.filename ) )
    def openUrl( self, url ):
        #http://www.oooforum.org/forum/viewtopic.phtml?t=35344
        properties = []
        properties.append( OfficeDocument._makeProperty( 'Hidden', True ) )
        properties.append( OfficeDocument._makeProperty( 'MacroExecutionMode', NEVER_EXECUTE ) )
        properties.append( OfficeDocument._makeProperty( 'ReadOnly', False ) )
        properties = tuple( properties )
        self.oodocument = self.openoffice.loadComponentFromURL( url, "_blank", 0, properties )
    def getFilename( self ):
        return self.filename
    def save( self ):
        self.oodocument.store()
        return self.getFilename()
    def saveAs( self, filename ):
        """Save the open office document to a new file, and possibly filetype. 
        The type of document is parsed out of the file extension of the filename given."""
        self.filename = os.path.abspath( filename )
        filename = uno.systemPathToFileUrl( self.filename )
        #filterlist: http://wiki.services.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
        exportFilter = self._getExportFilter( filename )
        props = exportFilter, 
        try:
            #if the filetype is a pdf this will crap-out
            self.oodocument.storeAsURL( filename, props )
        except:
            self.oodocument.storeToURL( filename, props )
        return self.getFilename()
    def saveTo( self, filename ):
        """Save the open office document to a new file, and possibly filetype. 
        The type of document is parsed out of the file extension of the filename given."""
        self.filename = os.path.abspath( filename )
        filename = uno.systemPathToFileUrl( self.filename )
        #filterlist: http://wiki.services.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
        exportFilter = self._getExportFilter( filename )
        props = exportFilter, 
        #storeToURL: #http://codesnippets.services.openoffice.org/Office/Office.ConvertDocuments.snip
        self.oodocument.storeToURL( filename, props )
        return self.getFilename()
    def close( self ):
        """Close the OpenOffice document"""
        self.oodocument.close( 1 )
    def _getExportFilter( self, filename ):
        """Automatically determine to output filter depending on the file extension"""
        #self._makeProperty( 'FilterName', 'writer_pdf_Export' )
        ext = OfficeDocument._getFileExtension( filename )
        for x in OfficeDocument.documentTypes:
            if self.oodocument.supportsService( x ):
                if x in OfficeDocument.exportFilters[ext].keys():
                    return OfficeDocument._makeProperty( 'FilterName', OfficeDocument.exportFilters[ext][x] )
        return None
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
        c = WriterDocument()
        c.open( inputFile )
        c.refresh()
        c.saveAs( outputFile )
        c.close()
    def _debug( self, unoobj, doPrint=False, type='both' ):
        """
        Print(or return) all of the method and property names for an uno object6
        Thanks to Carsten Haese for his insanely useful example fount at:
        http://bytes.com/topic/python/answers/641662-how-do-i-get-type-methods#post2545353
        """
        from com.sun.star.beans.MethodConcept import ALL as ALLMETHS
        from com.sun.star.beans.PropertyConcept import ALL as ALLPROPS
        from OfficeDocument import OfficeConnection
        ctx = OfficeConnection.context
        introspection = ctx.ServiceManager.createInstanceWithContext( "com.sun.star.beans.Introspection", ctx)
        access = introspection.inspect(unoobj)
        meths = access.getMethods(ALLMETHS)
        props = access.getProperties(ALLPROPS)
        if doPrint:
            if type == 'both' or type == 'methods':
                print "Object Methods:"
                for x in meths:
                    print "---- %s" % x.getName()
            if type == 'both' or type == 'properties':
                print "Object Properties:"
                for x in props:
                    print "---- %s" % x.Name
        #return [ x.getName() for x in meths ], [ x.Name for x in props ]
        return ( [x.getName() for x in meths], [x.Name for x in props] )
    def easydebug( self, obj ):
        meths, props = self._debug( obj )
        meths.sort()
        for x in meths:
            print '----- ' + x
        print ''
    def _debugMethod( self, unoobj, methodName ):
        from com.sun.star.beans.MethodConcept import ALL as ALLMETHS
        from com.sun.star.beans.PropertyConcept import ALL as ALLPROPS
        ctx = OfficeConnection.context
        introspection = ctx.ServiceManager.createInstanceWithContext( "com.sun.star.beans.Introspection", ctx)
        access = introspection.inspect(unoobj)
        method = access.getMethod( methodName, ALLMETHS )
        for x in method.getParameterInfos():
            print x
            print ""
        print ""
    def refresh( self, refreshIndexes=True ):
        """Refresh the OpenOffice document and (optionally) it's indexes"""
        try:
            self.oodocument.refresh()
        except:
            pass
        if refreshIndexes:
            #I needed the table of contents to automatically update in case page contents had changed
            try:
                #get all document indexes, eg. toc, or index
                oIndexes = self.oodocument.getDocumentIndexes()
                for x in range( 0, oIndexes.getCount() ):
                    oIndexes.getByIndex( x ).update()
            except:
                pass