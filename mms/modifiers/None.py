import sys, traceback
from mms.modifiers import modifiers, modifier
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
modifiers.modifierOrder.append( {'name':'None', 'order':100 } )
