python-openxml is a library to create and manipulate .docx and .pptx files.

The code draws heavily on the python-docx library created by Mike McCana at https://github.com/mikemaccana/python-docx/

python-openxml was written to support the Timetric data visualization platform (http://timetric.com)

For queries, please email Tom Scrace <tom.scrace@timetric.com>

Copyright Timetric Ltd., 2012.

---

To create a .docx file, suitable for use with Microsoft Word, do the following:

>>> from openxml.docx import Document
>>> d = Document.create()
>>> d.add_heading('Document heading')
>>> d.add_para('This is some text in the document')
>>> d.add_picture('image1.png')
>>> d.save('document.docx')

See the source code in docx.py for further details.

---

To create a .pptx file, suitable for use with Microsoft Powerpoint, do the following:


>>> from openxml.pptx import Document
>>> d = Document.create()
>>> s = d.add_slide()
>>> s.add_heading('Document heading')
>>> s.add_para('This is some text in the document')
>>> s.add_picture('image1.png')
>>> d.save('document.pptx')

See the source code in pptx.py for further details.
