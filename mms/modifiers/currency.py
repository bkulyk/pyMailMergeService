#from sys import path
import sys, os, traceback, locale
from mms.modifiers import modifiers, modifier
locale.setlocale( locale.LC_ALL, '' )

#try:
lc = mms.config.get( "mms", 'locale' )
if lc != '':
    locale.setlocale( locale.LC_ALL, tuple( mms.config.get( 'mms', 'locale' ).split( "." ) ) )
#except:
#    pass

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
            if x == '-1.0':
                x = 'Unlimited'
            elif x == '':
                x = ''
            else:
                x = '$' + locale.format("%d", int(round(float(x))), True, True)
        except Exception, e:
            pass
        return x
        
modifiers.modifierOrder.append( {'name':'currency', 'order':99 } )