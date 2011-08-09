import sys, os, traceback
from mms.modifiers import modifiers, modifier
import locale
locale.setlocale( locale.LC_ALL, '' )

try:
    from mms import mms
    lc = mms.config.get( "mms", 'locale' )
    if lc != '':
        locale.setlocale( locale.LC_ALL, tuple( mms.config.get( 'mms', 'locale' ).split( "." ) ) )
except:
    pass

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
                            result = '$' + locale.format("%.2f", float(x), True, True)
                            x = result
                    except Exception, e:
                        #traceback.print_exc(file=sys.stdout)
                        pass
                    document.searchAndReplaceFirst( key, x )
            else:
                try:
                    if param['value'] == '-1.0':
                        param['value'] = 'Unlimited'
                    else:
                        result = '$' + locale.format("%.2f", float(param['value']), True, True)
                        param['value'] = result
                except Exception, e:
                    #traceback.print_exc(file=sys.stdout)
                    pass
                document.searchAndReplace( key, param['value'] )
        except Exception, e:
            traceback.print_exc(file=sys.stdout)
modifiers.modifierOrder.append( {'name':'currencytwo', 'order':100 } )
