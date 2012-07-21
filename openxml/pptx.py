'''
Create and modify Open XML Presentation documents (.pptx, presentationML)
'''

import logging
from lxml import etree
try:
    from PIL import Image
except ImportError:
    import Image
import zipfile
import shutil
import re
import time
import os
from os.path import join
import tempfile
from namespaces import nsprefixes
from StringIO import StringIO

log = logging.getLogger(__name__)

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
template_dir = join(os.path.dirname(__file__),'pptx_template') # installed
if not os.path.isdir(template_dir):
    template_dir = join(os.path.dirname(__file__),'pptx_template') # dev

def relationshiplist():
    relationshiplist = [
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme','theme/theme1.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster','slideMasters/slideMaster1.xml'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide','slides/slide1.xml'],
    ]
    return relationshiplist
    
def contenttypes():
    # FIXME - doesn't quite work...read from string as temp hack...
    #types = makeelement('Types',nsprefix='ct')
    types = etree.fromstring('''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
    parts = {
        '/_rels/.rels':'application/vnd.openxmlformats-package.relationships+xml',
        '/ppt/_rels/presentation.xml.rels':'application/vnd.openxmlformats-package.relationships+xml',
        '/ppt/presentation.xml': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
        '/ppt/slides/_rels/slide1.xml.rels':'application/vnd.openxmlformats-package.relationships+xml',
        '/ppt/slides/slide1.xml': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
        '/ppt/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml',
        '/ppt/slideMasters/slideMaster1.xml': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        '/ppt/slideMasters/_rels/slideMaster1.xml.rels': 'application/vnd.openxmlformats-package.relationships+xml'
        }
    for i in range(1, 13):
        path1 = '/ppt/slideLayouts/slideLayout' + str(i) + '.xml'
        path2 = '/ppt/slideLayouts/_rels/slideLayout' + str(i) + '.xml.rels'
        parts[path1] = 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
        parts[path2] = 'application/vnd.openxmlformats-package.relationships+xml'
    
    for part in parts:
        types.append(makeelement('Override',nsprefix=None,attributes={'PartName':part,'ContentType':parts[part]}))
    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml','xml':'application/xml','jpeg':'image/jpeg','gif':'image/gif','png':'image/png'}
    for extension in filetypes:
        types.append(makeelement('Default',nsprefix=None,attributes={'Extension':extension,'ContentType':filetypes[extension]}))
    return types
    
def pptrelationships(relationshiplist):
    '''Generate a ppt relationships file'''
    # Default list of relationships
    # FIXME: using string hack instead of making element
    #relationships = makeelement('Relationships',nsprefix='pr')
    relationships = etree.fromstring(
    '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>'''
    )
    count = 0
    for relationship in relationshiplist:
        # Relationship IDs (rId) start at 1.
        relationships.append(makeelement('Relationship',attributes={'Id':'rId'+str(count+1),
        'Type':relationship[0],'Target':relationship[1]},nsprefix=None))
        count += 1
    return relationships

def makeelement(tagname,tagtext=None,nsprefix='p',attributes=None,attrnsprefix=None):
    '''Create an element & return it'''
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = nsprefixes[prefix]
        nsprefix = nsprefix[0] # FIXME: rest of code below expects a single prefix
    elif nsprefix:
        namespacemap = {nsprefix: nsprefixes[nsprefix]}
    else:
        # For when namespace = None
        nsprefix = 'p'
    namespace = '{'+nsprefixes[nsprefix]+'}'
    newelement = etree.Element(namespace+tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
            # Quick hack: it seems every element that has a 'w' nsprefix for its tag uses the same prefix for it's attributes
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
            attributenamespace = '{'+nsprefixes[attrnsprefix]+'}'

        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext
    return newelement
    
def picture(picname, slide_rels, picdescription='No Description', pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True,
            template=template_dir, align='center', scale=1):
    '''Take a relationshiplist, picture file name, and return a paragraph containing the image and an updated relationshiplist'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''
    # Copy the file into the media dir

    media_dir = join(template,'ppt','media')

    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
    new_picname = join(media_dir,os.path.basename(picname))
    shutil.copyfile(picname, new_picname)
    picname = new_picname
    
    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth,pixelheight = Image.open(picname).size[0:2]
    picname = os.path.basename(picname)
    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs
    emuperpixel = 12667
    width = str(int(pixelwidth * emuperpixel * scale))
    height = str(int(pixelheight * emuperpixel * scale))

    # Set relationship ID to the first available
    picid = len(slide_rels) + 1 
    picrelid = 'rId'+ str(picid)
    slide_rels.append([nsprefixes['i'], '../media/' + picname, str(picid)])

    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill')
    blipfill.append(makeelement('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
    stretch = makeelement('stretch',nsprefix='a')
    stretch.append(makeelement('fillRect',nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = makeelement('nvPicPr', nsprefix='p')
    cnvpr = makeelement('cNvPr', nsprefix='p',
                        attributes={'id': '37', 'name': 'BLAH'})
    nvpicpr.append(cnvpr)
    cnvpicpr = makeelement('cNvPicPr')
    nvpicpr.append(cnvpicpr)
    nvpicpr.append(makeelement('nvPr'))

    # 3. The Shape properties
    sppr = makeelement('spPr')
    xfrm = makeelement('xfrm',nsprefix='a')
    xfrm.append(makeelement('off',nsprefix='a',attributes={'x':'1405440','y':'1820520'}))
    xfrm.append(makeelement('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
    prstgeom = makeelement('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(makeelement('avLst',nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
    pic = makeelement('pic', nsprefix='p')
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    return slide_rels, pic
    
def savepptx(document, output, slides, media_files, pptrelationships,
                                    contenttypes=contenttypes(), template=template_dir):
    '''Save a modified document'''
    assert os.path.isdir(template)
    docxfile = zipfile.ZipFile(output,mode='w',compression=zipfile.ZIP_DEFLATED)

    # Serialize our trees into out zip file
    '''
    treesandfiles = {document:'ppt/presentation.xml',
                     contenttypes:'[Content_Types].xml',
                     pptrelationships:'ppt/_rels/presentation.xml.rels'}
    for tree in treesandfiles:
        log.info('Saving: '+treesandfiles[tree]    )
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr(treesandfiles[tree],treestring)
        '''
    for slide in slides:
        treestring = etree.tostring(slide.slide, pretty_print=True)
        parser = etree.XMLParser(ns_clean=True)
        tree = etree.parse(StringIO(treestring), parser)
        treestring = etree.tostring(tree, pretty_print=True)
        docxfile.writestr('ppt/slides/slide' + str(slide.number) + '.xml', treestring)
        rels_tree = etree.fromstring('''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>''')
        for rel in slide.relationships:
            rel_el = (etree.Element('Relationship'))
            rel_el.set('Id', 'rId' + rel[2])
            rel_el.set('Type', rel[0])
            rel_el.set('Target', rel[1])
            rels_tree.append(rel_el)
        rels_string = etree.tostring(rels_tree, pretty_print=True)
        docxfile.writestr('ppt/slides/_rels/slide' + str(slide.number) + '.xml.rels',
                                                                            rels_string)
    # Add & compress support files
    allowed = ['.xml', '.rels']
    for dirpath, dirnames, filenames in os.walk(template):
        for filename in filenames:
            ext = os.path.splitext(filename)[1]
            if ext in allowed or filename in allowed or filename in media_files:
                doc_file = os.path.join(dirpath, filename)
                archivename = doc_file[len(template)+1:]
                docxfile.write(doc_file, archivename)
    docxfile.close()
    return
    
def slide():
    sld = makeelement('sld', nsprefix=['p', 'r', 'a'])
    csld = makeelement('cSld')
    sptree = makeelement('spTree')
    
    nvgrpsppr = makeelement('nvGrpSpPr')
    cnvpr = makeelement('cNvPr', attributes={'id':'1', 'name':''})
    cnvgrpsppr = makeelement('cNvGrpSpPr')
    nvpr = makeelement('nvPr')
    nvgrpsppr.append(cnvpr)
    nvgrpsppr.append(cnvgrpsppr)
    nvgrpsppr.append(nvpr)
    
    sptree.append(nvgrpsppr)

    grpsppr = makeelement('grpSpPr')
    
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', attributes={'x':'0', 'y':'0'}, nsprefix='a'))
    xfrm.append(makeelement('ext', attributes={'cx':'0', 'cy':'0'}, nsprefix='a'))
    xfrm.append(makeelement('chOff', attributes={'x':'0', 'y':'0'}, nsprefix='a'))
    xfrm.append(makeelement('chExt', attributes={'cx':'0', 'cy':'0'}, nsprefix='a'))
    
    grpsppr.append(xfrm)
    
    sptree.append(grpsppr)
    csld.append(sptree)

    sld.append(csld)
    clrmapovr = makeelement('clrMapOvr')
    clrmapovr.append(makeelement('masterClrMapping', nsprefix='a'))
    sld.append(clrmapovr)
    return sld

def text_box(text):
    sp = makeelement('sp')
    nvsppr = makeelement('nvSpPr')
    nvsppr.append(makeelement('cNvPr', attributes={'id': '37', 'name': 'TextShape 1'}))
    nvsppr.append(makeelement('cNvSpPr', attributes={'txBox': '1'}))
    nvsppr.append(makeelement('nvPr'))
    sp.append(nvsppr)

    sppr = makeelement('spPr')

    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '332656'})) 
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': '9144000', 'cy': '1262160'}))
    sppr.append(xfrm)

    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
    sppr.append(prstgeom)

    sp.append(sppr)

    txbody = makeelement('txBody')
    txbody.append(makeelement('bodyPr', nsprefix='a', attributes={'anchor': 'ctr', 'bIns': '0', 'lIns': '0', 'rIns': '0', 'tIns': '0', 'wrap': 'none'}))
    p = makeelement('p', nsprefix='a')
    p.append(makeelement('pPr', nsprefix='a', attributes={'algn': 'ctr'}))
    r = makeelement('r', nsprefix='a')
    r.append(makeelement('rPr', nsprefix='a', attributes={'lang': 'en-GB'}))
    r.append(makeelement('t', nsprefix='a', tagtext=text)) # this is where the text goes.
    p.append(r)
    p.append(makeelement('endParaRPr', nsprefix='a'))
    txbody.append(p)
    
    sp.append(txbody)
    return sp

class Slide(object):
    def __init__(self):
        self.slide = slide()
        self.relationships = [
                [nsprefixes['sl'], '../slideLayouts/slideLayout2.xml', '1']
                ]
        self.number = None
        self.media_files = []
        return

    @classmethod
    def create(cls, template_dir):
        slide = cls()
        slide.template_dir = template_dir
        return slide

    def add_picture(self, picname, *args, **kwargs):
        extension = os.path.splitext(picname)[1]
        if extension not in ['.jpg', '.jpeg', '.png']:
            raise ValueError
        self.relationships, pic = picture(picname, slide_rels=self.relationships,
                                        template=self.template_dir, *args, **kwargs)
        self.slide.xpath('/p:sld/p:cSld/p:spTree', namespaces=nsprefixes)[0].append(pic)
        self.media_files.append(os.path.basename(picname))
        return

    def add_text_box(self, text):
        self.slide.xpath('/p:sld/p:cSld/p:spTree', namespaces=nsprefixes)[0].append(
                                                                          text_box(text))
        return

class Document(object):
    def __init__(self):
        self.relationshiplist = relationshiplist()
        self.slide_rels = [] # Each member of this list will be a list of relationships for a particular slide. Each relationship is itself a list, whose first member is the Type of the relationship (a namespace) and whose second member is the Target for the relationship.
        self.tmpdir = tempfile.mkdtemp()
        self.template_dir = os.path.join(self.tmpdir, 'template')
        shutil.copytree(template_dir, self.template_dir) # we copy our template files to a temp location
        return
    
    @classmethod
    def create(cls):
        doc = cls()
        doc.presentation = makeelement('presentation')
        master_id_list = makeelement('sldMasterIdLst')
        master_id_list.append(makeelement('sldMasterId', attributes={'id': '2147483648',                                                '{'+nsprefixes['r']+'}' + 'id':'rId2'}))
        doc.presentation.append(master_id_list)
        doc.presentation.append(makeelement('sldIdLst'))
        doc.presentation.append(makeelement('sldSz', attributes={'cx':'10080625',
                                                                 'cy':'7559675'}))
        doc.presentation.append(makeelement('notesSz', attributes={'cx':'7559675',
                                                                   'cy':'10691812'})) 
        doc.slides = []
        return doc

    def add_slide(self):
        slide = Slide.create(template_dir=self.template_dir)
        slide.number = len(self.slides) + 1
        self.slides.append(slide)
        slide_list = self.presentation.xpath('/p:presentation/p:sldIdLst',
                                                                namespaces=nsprefixes)[0]
        slide_list.append(makeelement('sldId',
            attributes={'id': str(256 + len(self.slides) - 1),
               '{'+nsprefixes['r']+'}' + 'id': 'rId' + str(3 + len(self.slides) - 1)}))
        return slide

    def save(self, filename, *args, **kwargs):
        media_files = []
        for slide in self.slides:
            media_files += slide.media_files
        suffix = '.pptx'
        if filename[-5:] != suffix: filename = filename + suffix
        return savepptx(document=self.presentation, slides=self.slides,
                        media_files=media_files, template=self.template_dir,
                        output=filename,
                        pptrelationships=pptrelationships(self.relationshiplist),
                        *args, **kwargs)

    def get_file_object(self, *args, **kwargs):
        '''Get the document as a file-like object.'''
        filedir = tempfile.mkdtemp()
        filepath = os.path.join(filedir, 'rendered_pptx.pptx')
        self.save(filename=filepath)
        f = open(filepath)
        shutil.rmtree(filedir)
        return f
        
    def get_as_string(self, *args, **kwargs):
        return self.get_file_object(*args, **kwargs).read()
 
    def close(self):
        shutil.rmtree(self.tmpdir)
