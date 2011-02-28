from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        print "here!!!"
        pass
modifiers.modifiers.modifierOrder.append( {'name':'if', 'order':10 } )