#from sys import path
import sys, os
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
                pass
            else:
                try:
                    if param['value'] == '-1.0':
                        param['value'] = 'Unlimited'
                    else:
                        result = '$' + locale.format("%.2f", float(param['value']), True, True)
                        param['value'] = result
                except Exception, e:
                    print sys.exc_info()
            document.searchAndReplace( key, param['value'] )
        except Exception, e:
            print sys.exc_info()
modifiers.modifiers.modifierOrder.append( {'name':'currencytwo', 'order':99 } )
