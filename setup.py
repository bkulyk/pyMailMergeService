from distutils.core import setup

setup( 
       name='MailMergeService',
       version='2.5', 
       description='OpenOffice/LibreOffice mail merge like document creation implemented over webservice',
       url="http://mailmergeservice.com",
       packages=['mms', 'mms.interfaces', 'mms.modifiers', 'mms.OfficeDocument'],
       scripts=['mmsd.py'],
       install_requires=[
            'pyPDF>=3',
            'uno',
            'cherrypy>=3',
            'soappy>0.12',
            'lxml>2.3' ] )