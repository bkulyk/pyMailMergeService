#from sys import path
import sys, os, traceback
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        if param['value'] == '1' :
            document.deleteColumn( key )
        else:
            document.searchAndReplace( key, '' )
        
modifiers.modifiers.modifierOrder.append( {'name':'deletecolumn', 'order':51 } )