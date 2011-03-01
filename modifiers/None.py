from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        if isinstance( param['value'], ( list, tuple ) ):
            for x in param['value']:
                document.searchAndReplaceFirst( "~%s~" % param['token'], x )
        else:
            #print "searching and replacing all %s --- value: %s" % ( param['token'], param['value'] )
            document.searchAndReplace( "~%s~" % param['token'], param['value'] )
modifiers.modifiers.modifierOrder.append( {'name':'multiparagraph', 'order':100 } )