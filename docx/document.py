import collections
import os
import zipfile
import shutil
import re

from docx import FILES_TO_IGNORE, NSPREFIXES
from docx.utils import make_element
from docx.meta import *


class DocxDocument(object):
    def __init__(self, template_file=None, template_dir=None):
        self.template_file = template_file
        self.template_dir = template_dir
        if self.template_file and os.path.isfile(self.template_file):
            self._init_from_file(self.template_file)
        else:
            self.template_file = None
            self.document = make_element("document")
            self.body = make_element("body")
            self.document.append(self.body)
            self.app_properties = AppProperties()
            self.core_properties = None
            self.word_relationships = WordRelationships()
            self.web_settings = WebSettings()
            self.content_types = ContentTypes()

    def _init_from_file(self, template_file):
        self.template_zip = zipfile.ZipFile(self.template_file, 'r', compression=zipfile.ZIP_DEFLATED)
        self.document = etree.fromstring(self.template_zip.read('word/document.xml'))
        self.word_relationships = WordRelationships(xml=self.template_zip.read('word/_rels/document.xml.rels'))
        #self.app_properties = AppProperties() # TODO: make available for manipulation
        #self.core_properties = CoreProperties() # TODO: make available for manipulation
        self.content_types = ContentTypes(xml=self.template_zip.read('[Content_Types].xml'))
        self.body = self.document.xpath('/w:document/w:body', namespaces=NSPREFIXES)[0]

    def search(self, search):
        '''Search a document for a regex, return success / fail result'''
        result = False
        searchre = re.compile(search)
        for element in self.document.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        result = element
        return result

    def replace(self, search, replace):
        '''Replace all occurences of string with a different string, return updated document'''
        searchre = re.compile(search)
        for element in self.document.iter():
            if element.tag == '{%s}t' % NSPREFIXES['w']: # t (text) elements
                if element.text:
                    if searchre.search(element.text):
                        if isinstance(replace, str) or isinstance(replace, unicode):
                            element.text = re.sub(search,replace,element.text)
                        else:
                            parent = element.getparent()
                            parent.replace(element, replace)
                            #element.addnext(element)

    def add(self, element, position=None):
        if position:
            # TODO: enable adding stuff to at specific points of the text.
            pass
        else:
            self.body.append(element)

    def get_text(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        paratextlist=[]   
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in self.document.iter():
            # Find p (paragraph) elements
            if element.tag == '{'+NSPREFIXES['w']+'}p':
                paralist.append(element)    
        # Since a single sentence might be spread over multiple text elements, iterate through each 
        # paragraph, appending all text (t) children to that paragraphs text.     
        for para in paralist:      
            paratext = u''  
            # Loop through each paragraph
            for element in para.iter():
                # Find t (text) elements
                if element.tag == '{'+NSPREFIXES['w']+'}t':
                    if element.text:
                        paratext = paratext+element.text
            # Add our completed paragraph text to the list of paragraph text    
            if not len(paratext) == 0:
                paratextlist.append(paratext)                    
        return paratextlist

    def append(self, element):
        self.document.append(element)

#    def _clean(self):
#        """ Perform misc cleaning operations on documents.
#            Returns cleaned document.
#        """
#        # Clean empty text and r tags
#        for t in ('t', 'r'):
#            rmlist = []
#            for element in self.document.iter():
#                if element.tag == '{%s}%s' % (NSPREFIXES['w'], t):
#                    if not element.text and not len(element):
#                        rmlist.append(element)
#            for element in rmlist:
#                element.getparent().remove(element)
#        return newdocument

    def _write_xml_files(self):
        # Serialize our trees into out zip file
        files = {
            self.core_properties: 'docProps/core.xml',
            self.app_properties: 'docProps/app.xml',
            self.content_types:'[Content_Types].xml',
            self.web_settings:'word/webSettings.xml',
            self.word_relationships:'word/_rels/document.xml.rels'
        }
        for f in files:
            treestring = etree.tostring(f._xml(), pretty_print=True)
            self.zip_file.writestr(files[f],treestring)

    def _copy_template_dir(self):
        """Copy a template document to our container."""
        for (dirpath, dirnames, filenames) in os.walk(self.template_dir):
            for filename in filenames:
                if filename in FILES_TO_IGNORE:
                    continue
                path = os.path.join(dirpath, filename)
                self.zip_file.write(path, os.path.relpath(path, self.template_dir))

    def _copy_template_file(self):
        """ Copy contents of template docx file into new docx file """
        for filename in self.template_zip.namelist():
            if os.path.basename(filename) in ['document.xml' , 'document.xml.rels', '[Content_Types].xml']:
                continue
            self.zip_file.writestr(filename, self.template_zip.read(filename))

    def _copy_media_files(self):
        for name, path in self.word_relationships.to_copy:
            out = 'word/media/' + name # IN DESPERATE NEED OF A FIX
            self.zip_file.write(path, out)

    def save(self, filename):
        '''Save a modified document'''
        self.zip_file = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)
        # TODO: determine what to do when template_file AND template_dir are specified
        if self.template_dir:
            self._write_xml_files()
            self._copy_template_dir()
        if self.template_file:
            self._copy_template_file()
            self.zip_file.writestr('word/_rels/document.xml.rels',
                                    etree.tostring(self.word_relationships._xml(),
                                    pretty_print=True, xml_declaration=True, encoding="utf-8"))
            self.zip_file.writestr('[Content_Types].xml',
                                    etree.tostring(self.content_types._xml(),
                                    pretty_print=True, xml_declaration=True, encoding="utf-8"))

        # Copying over any newly added media files.
        self._copy_media_files()
        # Adding the content file.
        self.zip_file.writestr('word/document.xml', etree.tostring(self.document, pretty_print=True, xml_declaration=True, encoding="utf-8"))
        
        # self.zip_file.close()
