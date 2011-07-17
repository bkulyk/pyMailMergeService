import sys, traceback
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        start = "~%s~" % param['token']
        end = "~end%s~" % param['token']
        value = "%s" % param['value']
        try:
            if value == '0':
                document.searchAndRemoveSection( start, end )
            else:
                document.searchAndReplace( start, '' )
                document.searchAndReplace( end, '' )
        except Exception, e:
            traceback.print_exc( file=sys.stdout )
modifiers.modifierOrder.append( { 'name':'if', 'order':10 } )
