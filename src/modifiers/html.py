from sys import path
import tempfile
path.append( '..' )
from modifiers import *
import modifiers, os
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        #save the html to a temporary html file
        fileHandle, fileName = tempfile.mkstemp( suffix="_pyMMS.html" )
        os.close( fileHandle )
        fileHandle = open( fileName, 'w' )
        fileHandle.write( param['value'] )
        fileHandle.close()
        #insert temporary document into the text document
        document.searchAndReplaceWithDocument( key, fileName )
        #remove the temporary document
        os.unlink( fileName )
modifiers.modifiers.modifierOrder.append( {'name':'html', 'order':30 } )