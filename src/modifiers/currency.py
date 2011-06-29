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
        try:
            if isinstance( param['value'], ( list, tuple ) ):
                for x in param['value']:
                    try:
                        if x == '-1.0':
                            x = 'Unlimited'
                        elif x == '':
			    x = ''
			else:
                            result = '$' + locale.format("%d", int(round(float(x))), True, True)
                            x = result
                    except Exception, e:
                        traceback.print_exc(file=sys.stdout)
                    document.searchAndReplaceFirst( key, x )
            else:
                try:
                    if param['value'] == '-1.0':
                        param['value'] = 'Unlimited'
                    else:
                        result = '$' + locale.format("%d", int(round(float(param['value']))), True, True)
                        param['value'] = result
                except Exception, e:
                    traceback.print_exc(file=sys.stdout)
	        document.searchAndReplace( key, param['value'] )
        except Exception, e:
            traceback.print_exc(file=sys.stdout)
modifiers.modifiers.modifierOrder.append( {'name':'currency', 'order':99 } )
