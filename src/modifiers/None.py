import sys, traceback
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        try:
            if isinstance( param['value'], ( list, tuple ) ):
                for x in param['value']:
                    document.searchAndReplaceFirst( key, x )
            else:
                count = document.searchAndReplace( key, param['value'] )
            document.drawSearchAndReplace( key, param['value'] )
        except:
            #key does not exist
            pass
modifiers.modifiers.modifierOrder.append( {'name':'None', 'order':100 } )
