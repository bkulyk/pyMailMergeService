#!/usr/bin/env python
import uno
from com.sun.star.beans import PropertyValue
#from com.sun.star.lang import XComponent
#from com.sun.star.frame import XStorable
from com.sun.star.io import IOException
import os
class OpenOfficeConnection:
    desktop = None
    host = "localhost"
    port = 8100
    @staticmethod
    def getConnection( **kwargs ):
        port = kwargs.get( 'port', OpenOfficeConnection.port )
        host = kwargs.get( 'host', OpenOfficeConnection.host )
        if OpenOfficeConnection.desktop is None:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext( 'com.sun.star.bridge.UnoUrlResolver', local )
            context = resolver.resolve( "uno:socket,host=%s,port=%s;urp;StarOffice.ComponentContext" % ( host, port ) )
            OpenOfficeConnection.desktop = context.ServiceManager.createInstanceWithContext( 'com.sun.star.frame.Desktop', context )
        return OpenOfficeConnection.desktop
class OpenOfficeDocument:
    openoffice = None
    oodocument = None
    def __init__(self):
        self.openoffice = OpenOfficeConnection.getConnection()
    def open( self, filename ):
        #http://www.oooforum.org/forum/viewtopic.phtml?t=35344
        properties = []
        properties.append( self._makeProperty( 'Hidden', True ) ) 
        properties = tuple( properties )
        self.oodocument = self.openoffice.loadComponentFromURL( uno.systemPathToFileUrl( os.path.abspath( filename ) ), "_blank", 0, properties )
    def refresh( self, refreshTOC=True ):
        try:
            self.oodocument.refresh()
        except:
            pass
        if refreshTOC:
            #I needed the table of contents to automatically update in case page contents had changed
            try:
                #get all document indexes, eg. toc, or index
                oIndexes = document.getDocumentIndexes()
                for x in range( 0, oIndexes.getCount() ):
                    oIndexes.getByIndex( x ).update()
            except:
                pass
    def saveAs( self, filename ):
        filename = uno.systemPathToFileUrl( os.path.abspath( filename ) )
        #filterlist: http://wiki.services.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
        props = tuple( [self._makeProperty( 'FilterName', 'writer_pdf_Export' )] ) 
        #storeToURL: #http://codesnippets.services.openoffice.org/Office/Office.ConvertDocuments.snip
        self.oodocument.storeToURL( filename, props )
    def close( self ):
        self.oodocument.close( 1 )
    def _makeProperty( self, key, value ):
        property = PropertyValue()
        property.Name = key
        property.Value = value
        return property
    @staticmethod
    def pdf( inputFile, outputFile ):
        c = OpenOfficeDocument()
        c.open( inputFile )
        c.refresh()
        c.saveAs( outputFile )
        c.close() 
#====================================================================        
if __name__ ==  '__main__':
    OpenOfficeDocument.pdf( "docs/invoice.odt", "invoice.pdf" )
    OpenOfficeDocument.pdf( "docs/simple_repeat_column.odt", "invoice2.pdf" )
