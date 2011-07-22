#from sys import path
import sys, os, traceback, locale
from mms.modifiers import modifiers, modifier
locale.setlocale(locale.LC_ALL, '')
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
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
            if x == '':
                x = ''
            else:
                x = locale.format("%d", int(round(float(x))), True, True)
        except Exception, e:
            pass
        return x
        
modifiers.modifierOrder.append( {'name':'number', 'order':101 } )