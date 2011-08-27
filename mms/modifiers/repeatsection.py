import sys
from mms.modifiers import modifiers, modifier
class __init__( modifier ):
    @staticmethod
    def process( document, param ):
        try:
            #the value should contain the number of times to repeat the section
            startkey = "~%s~" % param['token']
            endkey = "~end%s~" % param['token']
            count = int( param['value'] ) - 1
            #repeat the section of content wrapped by the start and end keys
            if document.searchAndDuplicate( startkey, endkey, count ) == 0:
                document.searchAndDuplicateInTable( startkey, endkey, count )
            #clean up (remove tokens from document)
            document.searchAndDelete( startkey )
            document.searchAndDelete( endkey )
        except Exception, e:
            print sys.exc_info()

modifiers.modifierOrder.append( {'name':'repeatsection', 'order':15 } )
