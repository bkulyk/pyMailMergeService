from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        pass
modifiers.modifiers.modifierOrder.append( {'name':'document', 'order':5 } )