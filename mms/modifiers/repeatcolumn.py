import sys
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = '~%s~' % param['token']
        count = 0
        if isinstance( param['value'], ( list, tuple ) ):
            document.duplicateColumn( key, len( param['value'] )-1 )
            for x in param['value']:
                y = document.searchAndReplaceFirst( key, x, True )
        else:
            document.searchAndReplaceFirst( key, param['value'] )

modifiers.modifierOrder.append( {'name':'repeatcolumn', 'order':25 } )
