import uno
class OfficeConnection:
    """Object to hold the connection to openoffice, so there is only ever one connection"""
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