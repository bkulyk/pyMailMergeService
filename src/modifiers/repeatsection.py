from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        #the value should contain the number of times to repeat the section
        startkey = "~%s~" % param['token']
        endkey = "~end%s~" % param['token']
        count = int( param['value'] ) - 1
        #repeat the section of content wrapped by the start and end keys
        if document.searchAndDuplicate( startkey, endkey, count ) == 0:
            document.searchAndDuplicateInTable( startkey, endkey, count )
        #clean up (remove tokens from document)
        document.searchAndDelete( startkey )
        document.searchAndDelete( endkey )
modifiers.modifiers.modifierOrder.append( {'name':'repeatsection', 'order':15 } )