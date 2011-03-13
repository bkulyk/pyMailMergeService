from sys import path
path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = '~%s~' % param['token']
        count = 0
        if isinstance( param['value'], ( list, tuple ) ):
            for x in param['value']:
                if count > 0: #first row already exists so we don't need to duplicate it
                    document.duplicateRow( key )
                count += 1
            for x in param['value']:
                document.searchAndReplaceFirst( key, x )
        else:
            document.searchAndReplaceFirst( key, param['value'] )
modifiers.modifiers.modifierOrder.append( {'name':'repeatrow', 'order':20 } )