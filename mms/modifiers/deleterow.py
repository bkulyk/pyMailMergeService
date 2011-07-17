import sys, os, traceback
from mms.modifiers import modifiers, modifier
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
modifiers.modifierOrder.append( {'name':'deleterow', 'order':50 } )