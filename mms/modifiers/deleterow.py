import sys, os, traceback
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        try:
            if param['value'] == '1' :
                more = 1
                i = 0
                while more == 1 and i < 10:
                    more = document.deleteRow( key )
                    i += 1
            else:
                document.searchAndReplace( key, '' )
        except:
            pass
modifiers.modifierOrder.append( {'name':'deleterow', 'order':50 } )