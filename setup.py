#!/usr/bin/env python

from distutils.core import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

setup(name='openxml',
      version='0.1',
      requires=['lxml'],
      description='Create .pptx and .docx files from Python',
      author='Mike MacCana',
      author_email='python.docx@librelist.com',
      url='http://github.com/mikemaccana/python-docx',
      py_modules=['openxml', 'pptx', 'docx'],
      data_files=[
          ('docx_template/_rels', glob('docx_template/_rels/.*')),
          ('docx_template/docProps', glob('docx_template/docProps/*.*')),
          ('docx_template/word', glob('docx_template/word/*.xml')),
          ('docx_template/word/theme', glob('docx_template/word/theme/*.*')),
          ('pptx_template/_rels', glob('pptx_template/_rels/.*')),
          ('pptx_template/ppt', glob('pptx_template/ppt/*.xml')),

          ],
      )
