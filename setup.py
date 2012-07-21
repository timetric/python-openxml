#!/usr/bin/env python

from distutils.core import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='python-openxml',
      version='0.1',
      requires=['lxml'],
      description='Create .pptx and .docx files from Python',
      author='Tom Scrace',
      author_email='tom.scrace@timetric.com',
      url='http://github.com/timetric/python-openxml',
      packages=['openxml'],
      data_files=[
          ('openxml/docx_template/_rels', glob('docx_template/_rels/.*')),
          ('openxml/docx_template/docProps', glob('docx_template/docProps/*.*')),
          ('openxml/docx_template/word', glob('docx_template/word/*.xml')),
          ('openxml/docx_template/word/theme', glob('docx_template/word/theme/*.*')),
          ('openxml/pptx_template/_rels', glob('pptx_template/_rels/.*')),
          ('openxml/pptx_template/ppt', glob('pptx_template/ppt/*.xml')),
          ],
      )
