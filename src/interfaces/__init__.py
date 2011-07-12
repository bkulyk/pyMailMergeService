import os, sys
sys.path.append( ".." )
from lib.pyMailMerge import pyMailMerge
import simplejson as json
import tempfile
import shutil
from com.sun.star.task import ErrorCodeIOException
import traceback
import StringIO
class base( object ):
    outputDir = stubsDir = documentBase = "./docs/"
    def __init__( self, options={} ):
        if 'outputDir' in options:
            self.outputDir = options['outputDir'] + "/"
        if 'stubsDir' in options:
            self.stubsDir = options['stubsDir'] + "/"
            
    def uploadConvert( self, params='', odt='', type='pdf' ):
        #write temporary file
        file, fileName = tempfile.mkstemp( suffix="_pyMMS.odt" )
        os.close( file )
        file = open( fileName, 'w' )
        file.write( odt )
        file.close()
        #do conversion
        try:
            mms = pyMailMerge( fileName )
            content = mms.convert( params, type )
            os.unlink( fileName )
            return content
        except:
            number = '?'
            message = "unknown exception"
        return self.__errorXML( number, message )
    
    def pdf( self, params='', odt='' ):
        return self.convert(params, odt, 'pdf')
    
    def convert( self, params='', odt='', type='pdf', resave=False, saveExport=False ):
        try:
            fileName = os.path.abspath( self.outputDir + odt )
            mms = pyMailMerge( fileName )
            return mms.convert( params, type, resave, saveExport )
        except:
            number = '?'
            message = "unknown exception"
        return self.__errorXML( number, message )
    
    def calculator( self, params='', ods='', format='json', outputFile=None ):
#        try:
            fileName = os.path.abspath( self.stubsDir + ods )
            mms = pyMailMerge( fileName, 'ods' )
            if outputFile is not None:
                data = mms.calculator( params, self.outputDir + outputFile )
            else:
                data = mms.calculator( params )
            #todo xml output
            print data
            data = json.dumps( data )
            return data
#        except:
#            number = '?'
#            message = 'unknown exeption in calculator'
#            message += "\n\n"
#            tmp = StringIO.StringIO()
#            traceback.print_exc( file=tmp )
#            message += tmp.getvalue()
#            
#        return self.__errorXML( number, message )
    
    def batchConvert( self, batch=None, outputType='pdf', outputFilePath=False ):
#        try:
            batch = json.loads( batch )
#        except:
#            pass
# 
#        try:       
            #process tokens for each document individually
            files = []
            i=0
            for doc, params in batch:
                mm = pyMailMerge( self.stubsDir + doc )
                outputFile = mm.convertFile( params, 'odt' )
                files.append( outputFile )
                
                #some stuff for debugging.
#                x = open( "/var/www/mms_docs/output/%s_%s.xml" % ( doc, i ), 'w' )
#                x.write( params )
#                x.close()
#                i+=1
                
            #open first file
            mm = pyMailMerge( files[0] )
            
            #append every other file to the first
            for x in files[1:]:
                mm.joinDocumentToEnd( x )
                
            if outputFilePath is not False:
                #convert file and get path
                outputFile = mm.convertFile( None, outputType )
                #copy file to putput dir
                outputFilePath = os.path.abspath( os.path.join( self.outputDir, outputFilePath ) )
                shutil.copy( outputFile, outputFilePath )
                #return the filename
                content = outputFilePath
                #cleanout
                os.unlink( outputFile )
            else:
                content = mm.convert( None, 'odt' )
            
            for x in files:
                os.unlink( x )
            
            return content
#        except:
#            number = "?"
#            message = 'unknown exception in batchConvert'
#            message += "\n\n"
#            tmp = StringIO.StringIO()
#            traceback.print_exc( file=tmp )
#            message += tmp.getvalue()
#            return self.__errorXML( number, message )        
    
    def batchpdf( self, batch ):
        from pyPdf import PdfFileWriter, PdfFileReader
        output = PdfFileWriter()
        #convert each document into a pdf
        for x in batch:
            odtname, tokens = x
            pdfpath = self.pdf( odtname, tokens )
            pdf = PdfFileReader( file( pdfpath, 'rb' ) )
            #add each page of the new pdf to the big pdf
            for x in range( 0, pdf.getNumPages() ):
                output.addPage( pdf.getPage( x ) )
        #write the merged pdf to disk
        outfile = self._getTempFile('.pdf')
        outputStream = file( outfile, 'wb' )
        output.write( outputStream )
        outputStream.close()
        #read all the pdf back so it can be sent to the client
        readin = open( outfile, 'r' )
        contents = readin.read()
        return contents
    
    def getTokens( self, odt='', format='json' ):
        try:
            path = os.path.abspath( self.stubsDir + odt )
            mms = pyMailMerge( path )
            tokens = mms.getTokens()
            if format=='xml':
                xml = """<?xml version="1.0" encoding="UTF-8"?><tokens>"""
                for x in tokens:
                    xml += "<token>%s</token>" % x
                xml += "</tokens>"
                return xml
            else:
                return json.dumps( tokens )
        except:
            return self.__errorXML( '?', 'could not get tokens' )
    def getNamedRanges( self, ods='', format='json' ):
        path = os.path.abspath( self.stubsDir + odt )
        mms = pyMailMerge( path, 'ods' )
        names = mms.getNamedRanges()
        if format=='xml':
            xml = """<?xml version="1.0" encoding="UTF-8"?><namedranges>"""
            for x in names:
                xml += "<namedrange>%s</namedrange>" % x
            xml += "</namedranges>"
            return xml
        else:
            return json.dumps( names )
        
    def __errorXML( self, number, message ):
        return """
        <?xml version="1.0" encoding="UTF-8"?>
        <errors>
            <error>
                <number>%s</number>
                <message><![CDATA[%s]]></message>
            </error>
        </errors>""" % (number, message)
