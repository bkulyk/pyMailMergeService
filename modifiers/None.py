from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        if isinstance( param['value'], ( list, tuple ) ):
            for x in param['value']:
                document.searchAndReplaceFirst( key, x )
        else:
            #print "searching and replacing all %s --- value: %s" % ( param['token'], param['value'] )
            count = document.searchAndReplace( key, param['value'] )
            if count == 0:
                document.drawSearchAndReplace( key, param['value'] )
modifiers.modifiers.modifierOrder.append( {'name':'None', 'order':100 } )