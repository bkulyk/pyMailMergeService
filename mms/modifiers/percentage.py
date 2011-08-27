#from sys import path
import sys, os, traceback, locale
from mms.modifiers import modifiers, modifier

class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        debug = False
        try:
            if isinstance( param['value'], ( list, tuple ) ):
                for x in param['value']:
                  x = __init__.cleanVal(x)
                  document.searchAndReplaceFirst( key, x )
            else:
                x = param['value']
                x = __init__.cleanVal(x)
                count = document.searchAndReplace( key, x )
        except Exception, e:
            traceback.print_exc(file=sys.stdout)
    @staticmethod
    def cleanVal( x ):
        try:
            x = "%d" % int( round( float(x) ) )
            x += "%"
        except Exception, e:
            x = ''
        return x
            
modifiers.modifierOrder.append( {'name':'percentage', 'order':97 } )
