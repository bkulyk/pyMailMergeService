from sys import path
import tempfile
path.append( '..' )
from modifiers import *
import modifiers, os
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        fileHandle, fileName = tempfile.mkstemp( suffix="_pyMMS.html" )
        os.close( fileHandle )
        fileHandle = open( fileName, 'w' )
        fileHandle.write( param['value'] )
        fileHandle.close()
        document.searchAndReplaceWithDocument( key, fileName )
modifiers.modifiers.modifierOrder.append( {'name':'html', 'order':30 } )