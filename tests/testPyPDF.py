import unittest
from pyPdf import PdfFileWriter, PdfFileReader
import os

class testPyPdf( unittest.TestCase ):
    def setUp(self):
        pass
    def test_concat_pdf_files( self ):
        try:
            os.unlink( r"docs/c.pdf" )
        except:
            pass
        self.assertTrue( True )
        input_a = PdfFileReader( file( r"docs/a.pdf", 'rb' ) )
        input_b = PdfFileReader( file( r"docs/b.pdf", 'rb' ) )

        output = PdfFileWriter()

        for x in range( 0, input_a.getNumPages() ):
            output.addPage( input_a.getPage( x ) )
        for x in range( 0, input_b.getNumPages() ):
            output.addPage( input_b.getPage( x ) )

        outputStream = file( r"docs/c.pdf", 'wb' )
        output.write( outputStream )
        outputStream.close()
        
        count = input_a.getNumPages() + input_b.getNumPages()
        
        check = PdfFileReader( file( r"docs/c.pdf", 'rb' ) )
        self.assertEqual( count, check.getNumPages() )
        os.unlink( r"docs/c.pdf" )
if __name__ == '__main__':
    unittest.main()        