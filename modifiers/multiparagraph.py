from sys import path
path.append( '..' )
import html
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        html.__init__.process( document, param )
modifiers.modifiers.modifierOrder.append( { 'name':'multiparagraph', 'order':35 } )