import sys, traceback
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        if isinstance( param['value'], ( list, tuple ) ):
            for x in param['value']:
                if x == '1':
                    x = 'Yes'
                else:
                    x = 'No'
                document.searchAndReplaceFirst( key, x )
        else:
            if param['value'] == '1':
                value = 'Yes'
            else:
                value = 'No'
            count = document.searchAndReplace( key, value )
            try:
                document.drawSearchAndReplace( key, value )
            except:
                pass
modifiers.modifiers.modifierOrder.append( {'name':'yesno', 'order':32 } )