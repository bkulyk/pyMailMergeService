#from sys import path
import sys, os, traceback
sys.path.append( '..' )
from modifiers import *
import modifiers
import locale
locale.setlocale(locale.LC_ALL, '')
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
                x = __init__.cleanVal( param['value'] )
                document.searchAndReplace( key, x )
        except Exception, e:
            traceback.print_exc(file=sys.stdout)
    @staticmethod
    def cleanVal( x ):
        try:
            x = "%.2f%%" % float( x )
        except Exception, e:
            x = ''
        return x
modifiers.modifiers.modifierOrder.append( {'name':'percentagetwo', 'order':98 } )
