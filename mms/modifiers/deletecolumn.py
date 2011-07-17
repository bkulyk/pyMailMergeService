import sys, os, traceback
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        if param['value'] == '1' :
            document.deleteColumn( key )
        else:
            document.searchAndReplace( key, '' )
        
modifiers.modifierOrder.append( {'name':'deletecolumn', 'order':51 } )
