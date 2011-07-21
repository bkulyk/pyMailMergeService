import sys, os, traceback
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        if param['value'] == '1' :
            more = 1
            i = 0
            while more == 1 and i < 10:
                more = document.deleteColumn( key )
                i += 1
        else:
            document.searchAndReplace( key, '' )
        
modifiers.modifierOrder.append( {'name':'deletecolumn', 'order':51 } )
