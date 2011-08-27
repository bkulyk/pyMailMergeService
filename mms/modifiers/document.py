from sys import path
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        key = "~%s~" % param['token']
        fileName = param['value']
        document.searchAndReplaceWithDocument( key, fileName )
modifiers.modifierOrder.append( {'name':'document', 'order':5 } )