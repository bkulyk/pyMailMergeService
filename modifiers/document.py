from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        fileName = param['value']
        document.searchAndReplaceWithDocument( key, fileName )
modifiers.modifiers.modifierOrder.append( {'name':'document', 'order':5 } )