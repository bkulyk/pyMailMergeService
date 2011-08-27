import ConfigParser, os

class Config:
    config = None

    @staticmethod
    def parseConfig():
        config = ConfigParser.ConfigParser()
        #include the default config file.
        defaultPath = os.path.join( os.path.dirname( __file__ ), 'defaults.cfg' )
        #override with file in install path, then override with file in users home dir if present
        config.read( [ defaultPath, 'mms.cfg', os.path.expanduser("~/.mms.cfg") ] )
        
        Config.config = config
        return config

def getConfig():
    if Config.config is None:
        return Config.parseConfig()
    else:
        return Config.config