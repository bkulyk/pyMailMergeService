import sys
from mms.OfficeDocument import OfficeDocument
from mms.modifiers import modifiers
import mms.exceptions as mms_exceptions
from lxml import etree              #for parsing xml parameters
import re                           #regular expressions
import operator                     #using for sorting the params
import os                           #removing the temp files
import tempfile                     #create temp files for writing xml and output
import shutil, ConfigParser, logging
from logging import handlers
class pyMailMerge:
    document = None
    documentPath = None
    origionalPath = None
    log = None
    
    documentsOpen = { } #should allow me to keep track of how many document are supposed to be open at a time. 
    
    def __init__( self, odt='', type='odt', log=None ):
        #reuse logging object
        self.log = mms.logger
        
        self.log.debug( "create %s document" % type )
        self.document = OfficeDocument.createDocument( type )
        if odt != '':
            try:
                #copy filt to temporary document, becasuse two webservice users using the same file causes problems
                self.origionalPath = odt 
                self.documentPath = self._getTempFile( type )
                shutil.copyfile( odt, self.documentPath )
                self.log.debug( 'copied file from: %s to: %s' % (odt, self.documentPath) )
                
                self.document.open( self.documentPath )
            except:
                self.log.error( "Stub '%s' could not be found" % odt )
                raise mms_exceptions.StubNotFound( "Stub '%s' could not be found" % odt )
            
            pyMailMerge.documentsOpen[ self.documentPath ] = self.documentPath
            
            self.log.info( "opened office document: %s" % self.documentPath )
            
    def __del__(self):
        self.document.close()
        
        #remove entry from dict of currently open documents
        if self.documentPath in pyMailMerge.documentsOpen.keys():
            del( pyMailMerge.documentsOpen[ self.documentPath ] )
            
        try:
            if self.documentPath is not None:
                if os.path.exists( self.documentPath ):
                    os.unlink( self.documentPath )
        except:
            #I don't really care about any errors at this point...
            #I should have caught them when opening or connecting
            pass
    def getTokens( self ):
        tokens = self.document.re_match( r"~[a-zA-Z0-9\_\|\:\.\?]+~" )
        return dict( map( lambda i: ( i, 1 ), tokens ) ).keys()
    
    def joinDocumentToEnd( self, fileName ):
        return self.document.addDocumentToEnd( fileName )

    def getNamedRanges( self ):
        return self.document.getNamedRanges()

    def calculator( self, xml, outputFile=None ):
        if outputFile is not None:
            self.document.saveAs( outputFile )
        try:
            #process like normal mail merge 
            params = pyMailMerge._readParamsFromXML( xml )
            self._process( params )
        except:
            print 'no tokens'
            pass
        
        xml = etree.XML( xml )
        positivevalues = [ '1', 1, 'true', True, 'True', 'yes', 'Yes' ]
        
        #for each input namedrange set the values
        for x in xml.xpath( "//input/namedranges/namedrange" ):
            
            #get the name, could be attribute or sub element
            name = x.get( 'name' )
            if name is None:
                name = x.findall( "name" )[0].text
            values = []
            for value in x.findall( 'value' ):
                #check to see if the value is supposed to be a number or a string
                rangenumeric = "%s" % x.get( 'numeric' )
                valuenumeric = "%s" % value.get( 'numeric' )
                if rangenumeric in positivevalues or valuenumeric in positivevalues:
                    values.append( pyMailMerge.toNumber( value.text ) )
                else:
                    values.append( value.text )
            
            #set the range values
            self.document.setNamedRangeStrings( name, values )
        
        self.document.refresh()
        
        data = {}    
        #extraced the data for the named ranges requested in the 'output' node
        for x in xml.xpath( "//output/namedrange" ):
            try:
                limit = int( x.get('limit') )
            except:
                limit = None
            try:
                data[ x.text ] = self.document.getNamedRangeStrings( x.text, limit )
            except:
                #would mean that the range does not exist
                data[ x.text ] = 'error range %s does not exist' % x.text 
        
        if outputFile is not None:
            self.document.save()
            self.document.close()
            try:
                os.chmod( outputFile, 0777 )
            except:
                pass
        
        return data
    
    def convert( self, params, type='pdf' ):
        try:
            out = self.convertFile( params, type  )
        except Exception, e:
            print sys.exc_info()

        #read contents
        if out is not None:
            file = open( out, 'r' )
            x = file.read()
            file.close()
            
            #clean up
            os.unlink( out )
        return x
    
    def convertFile(self, params, type='pdf' ):
        if params is not None and params != '':
            params = pyMailMerge._readParamsFromXML( params )
    
            #process params
            self._process(params)

        #get temporary out put file
        out = self._getTempFile( '.%s' % type )
        self.document.refresh()
        self.document.saveAs( out )
        
        return out

    def _process( self, params ):
        params = pyMailMerge._sortParams(params)
        for param in params:
            #each module has been set as a property of the modifiers class
            mod = getattr( modifiers, "mod_%s" % param['modifier'] )
            #so, all the modifiers have been loaded as class names called __init__ within their 
            #respective modules, and each modifier should have a method called process
#            try:
            mod.__init__().process( self.document, param )
#            except:
#                pass
    def _getTempFile( self, extension='' ):
        '''
        This is the replacement for os.tmpnam, it should not be suceptable to the
        same symlink injection problems.  I wrapped this simple command in a function 
        because I don't want to use the file handle it creates. There maybe a more efficient
        way to get a tmp file name, but I havn't found one yet, so I'm using this.
        @param extension:  the file extension that the temp file should have.
        @return: str the temporary filename  
        '''
        file = tempfile.mkstemp( suffix="_pyMMS%s" % extension )
        #I am not going to use this file handle because I found in my research,
        #that I need to open it with the codec module to ensure it's open as utf-8 
        #file instead of whatever the default is
        os.close( file[0] )
        return file[1]
    @staticmethod
    def _readParamsFromXML(xml):
        xml = etree.XML( xml )
        tokens = xml.xpath( "//tokens/token" )
        params = []
        for x in tokens:
            tokenname = x.findall( "name" )[0].text
            tokenvalues = x.findall( "value" )
            if len( tokenvalues ) == 1:
                values = tokenvalues[0].text
            else:
                values = []
                for x in tokenvalues:
                    if x.text is None:
                        values.append ( "" )
                    else:
                        values.append( x.text )
            params.append( { 'token':tokenname, 'value':values } )
        return params
    @staticmethod
    def _readNamedRangesFromXML( xml ):
        if isinstance( xml, ( list, tuple ) ):
            #not actually xml
            return xml
        data = []
        xml = etree.XML( xml )
        namedRanges = xml.xpath( "//namedranges/name" )
        for x in namedRanges:
            data.append( x.text )
        return data
    @staticmethod
    def _sortParams(params):
        #break up by token, so they can be sorted later 
        sorted = { 'None':[] }
        for x in modifiers.modifierOrder:
            sorted[ x['name'] ] = []
        #now place each key in the right spot in sorted
        for x in params:
            result = re.match( r"(?P<modifier>^[0-9A-Za-z\._]+\|)*(?P<token>[0-9A-Za-z\._\:\|]+)", x['token'] )
            mod = "%s" % result.group( 'modifier' )
            mod = mod.replace( '|', '' )
            #add the modifier to the dictionary, this will save some effor later
            x['modifier'] = mod
            sorted[ mod ].append( x )
        #so far they are not actually sorted, just divided up, now lets put them all back into one list in the right order
        all = []
        for x in modifiers.modifierOrder:
            all.extend( sorted[ x['name'] ] )
        return all
    @staticmethod 
    def toNumber( number ):
        try:
            return float( number )
        except:
            pass
        try:
            return int( number )
        except:
            pass
        return number
    
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

class mms:
    config = None
    logger = None
    
    def __init__(self):
        #parse a config file
        self.config = mms.parseConfig()
        self.logger = mms.getLogger( self.config )
        
        interface = self.config.get( 'mms', 'interface' )
        #initiate one of the webservice interfaces, REST is going to be the default
        if interface == 'soap':
            from mms.interfaces.soap import soap
            soap.run( self.config, self.logger )
        elif interface == 'rest':
            from mms.interfaces.rest import rest
            rest.run( self.config, self.logger )
    
    @staticmethod
    def parseConfig():
        config = ConfigParser.ConfigParser()
        #include the default config file.
        defaultPath = os.path.join( os.path.dirname( __file__ ), 'defaults.cfg' )
        #override with file in install path, then override with file in users home dir if present
        config.read( [ defaultPath, 'mms.cfg', os.path.expanduser("~/.mms.cfg") ] )
        return config
    
    @staticmethod
    def getLogger( config ):
        enabled = config.get( "logging", 'enabled' ) in ( 'True', 'true' )
        if enabled:
            x = logging.getLogger( 'mms' )
            level = config.get( "logging", "level" ).upper()
            level = getattr( logging, level )
            x.setLevel( level )
            maxbyte = int( config.get( "logging", "maxbytes" ) )
            path = config.get( "logging", 'path' )
            handler = logging.handlers.RotatingFileHandler( path, maxBytes=maxbyte )
            formatter = logging.Formatter( "%(asctime)s - %(name)s - %(levelname)s - %(message)s" )
            handler.setFormatter( formatter )
            x.addHandler( handler )
            return x
        else:
            return NoLogging()
        
class NoLogging:
    def info(self, msg):
        pass
    def error(self, msg):
        pass
    def exception(self,msg):
        pass
    def debug(self,msg):
        pass
