import os, sys
sys.path.append( ".." )
from lib.pyMailMerge import pyMailMerge
import simplejson as json
import tempfile
from com.sun.star.task import ErrorCodeIOException
class base( object ):
    outputDir = stubsDir = documentBase = "../tests/docs/"
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
    def joinDocuments( self, odt, fileNames, addPageNumbers=False, appendToExistingDoc=False ):
        try:
            outputFileName = os.path.abspath( self.outputDir + odt )
            mms = pyMailMerge()

            if appendToExistingDoc:
                mms.document.open( outputFileName )
            else:
                mms.document.createNew()

            for x in fileNames.split( ':' ):
                inputFileName = os.path.abspath( self.stubsDir + x )
                mms.joinDocumentToEnd( inputFileName )
            if addPageNumbers != False:
                mms.document.applyPageNumberFooterToDefaultStyle()

            if appendToExistingDoc:
                mms.document.save()
            else:
                mms.document.saveAs( outputFileName )

            mms.document.close()
            return "Ok"
        except ErrorCodeIOException:
            number = "500"
            message = "Could not write completed file"
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
        #try:
            path = os.path.abspath( self.outputDir + odt )
            print path
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
        #except:
        #    return self.__errorXML( '?', 'could not get tokens' )
    def __errorXML( self, number, message ):
        return """
        <?xml version="1.0" encoding="UTF-8"?>
        <errors>
            <error>
                <number>%s</number>
                <message>%s</message>
            </error>
        </errors>""" % (number, message)
