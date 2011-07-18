#!/usr/bin/env python
from distutils.core import setup

setup( 
       name='MailMergeService',
       version='2.5', 
       description='OpenOffice/LibreOffice mail merge like document creation implemented over webservice',
       url="http://mailmergeservice.com",
       packages=['mms', 'mms.interfaces', 'mms.modifiers', 'mms.OfficeDocument'],
       scripts=['mms/mmsd.py'],
       requires=[
            'pyPDF (>=3.0)',
            'uno',
            'cherrypy (>=3.0)',
            'soappy (>0.12)',
            'lxml (>2.3)' ] )
