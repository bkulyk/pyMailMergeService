from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, key, value ):
        pass
modifiers.modifiers.modifierOrder.append( {'name':'html', 'order':30 } )