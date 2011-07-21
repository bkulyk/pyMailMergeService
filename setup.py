#!/usr/bin/env python
from setuptools import setup

setup( 
       name='MailMergeService',
       version='2.5', 
       description='OpenOffice/LibreOffice mail merge like document creation implemented over webservice',
       url="http://mailmergeservice.com",
       packages=['mms', 'mms.interfaces', 'mms.modifiers', 'mms.OfficeDocument'],
       scripts=['mmsd.py'],
       include_package_data=True,
       package_data={
                "": [ "*" ] #all files in all modules
            },
       install_requires=[
#                'pyPDF>=3.0',
#                'uno',
                'cherrypy>=3.0',
                'soappy>0.12',
#                'lxml>2.3' 
            ] )
