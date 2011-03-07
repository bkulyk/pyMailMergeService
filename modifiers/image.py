from sys import path
import base64, os
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        image = base64.b64decode( param['value'] )
        #save file to a temporary file.
        fileHandle, fileName = tempfile.mkstemp( suffix="_pyMMS.jpg" )
        os.close( fileHandle )
        fileHandle = open( fileName, 'w' )
        fileHandle.write( image )
        fileHandle.close()
        #@todo insert image into the text document
        
        #remove temporary file
        os.unlink( fileName )
modifiers.modifiers.modifierOrder.append( {'name':'image', 'order':40 } )