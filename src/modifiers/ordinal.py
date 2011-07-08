#from sys import path
import sys, os
sys.path.append( '..' )
from modifiers import *
import modifiers
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        try:
            if isinstance( param['value'], ( list, tuple ) ):
                pass
            else:
                try:
                    val = int(round(float(param['value'])))
                    if 10 <= val % 100 < 20:
                        param['value'] = str(val) + 'th'
                    else:
                        param['value'] = str(val) + {1 : 'st', 2 : 'nd', 3 : 'rd'}.get(val % 10, "th")
                except Exception, e:
                    pass #print sys.exc_info()
            document.searchAndReplace( key, param['value'] )
        except Exception, e:
            print sys.exc_info()
modifiers.modifiers.modifierOrder.append( {'name':'ordinal', 'order':97 } )
