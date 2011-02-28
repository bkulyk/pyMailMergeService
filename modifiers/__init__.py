import os, operator
class modifiers:
    modifierOrder = []
    @staticmethod
    def sortModifiers():
        #make sure the modifiers are in the right order
        modifiers.modifierOrder.sort(key=operator.itemgetter('order'))
    @staticmethod
    def init():
        modifiers.sortModifiers()
class modifier:
    pass
#import all modules in this folder. This could be a bad idea. If somebody has a better way, let me know.
for module in os.listdir( os.path.dirname( __file__ ) ):
    if module != '__init__.py' and module[-3:] == '.py':
        __import__( module[:-3], locals(), globals() )
        if module[:-3] in globals():
            setattr( modifiers, 'mod_%s' % module[:-3], globals()[ module[:-3] ] )
del module
modifiers.init()