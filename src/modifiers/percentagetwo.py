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
                    param['value'] = "%.2f%%" % float(param['value'])
                except Exception, e:
                    print sys.exc_info()
            document.searchAndReplace( key, param['value'] )
        except Exception, e:
            print sys.exc_info()
modifiers.modifiers.modifierOrder.append( {'name':'percentagetwo', 'order':98 } )
