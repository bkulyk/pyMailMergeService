from sys import path
import html
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        html.__init__.process( document, param )
modifiers.modifierOrder.append( { 'name':'multiparagraph', 'order':35 } )