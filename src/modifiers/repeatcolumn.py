import sys
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = '~%s~' % param['token']
        count = 0
        try:
            if isinstance( param['value'], ( list, tuple ) ):
                for x in param['value']:
                    if count > 0: #first column already exists so we don't need to duplicate it
                        document.duplicateColumn( key )
                    count += 1
                for x in param['value']:
                    document.searchAndReplaceFirst( key, x )
            else:
                document.searchAndReplaceFirst( key, param['value'] )
        except Exception, e:
            print sys.exc_info()
modifiers.modifiers.modifierOrder.append( {'name':'repeatcolumn', 'order':25 } )
