#from sys import path
import sys, os, traceback
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        try:
            if param['value'] == '1' :
                document.deleteRow( key )
            else:
                document.searchAndReplace( key, '' )
        except:
            pass
modifiers.modifiers.modifierOrder.append( {'name':'deleterow', 'order':50 } )