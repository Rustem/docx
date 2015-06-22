#!/usr/bin/env python2.6
# -*- coding: utf-8 -*-
'''
Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

Part of Python's docx module - http://github.com/mikemaccana/python-docx
See LICENSE for licensing information.
'''

from lxml import etree
try:
    from PIL import Image
except ImportError:
    # BBB for broken PIL installations
    import Image
import zipfile
import shutil
import re
import os
from os.path import join

# Record template directory's location which is just 'template' for a docx
# developer or 'site-packages/docx-template' if you have installed docx
#TEMPLATE_DIR = join(os.path.dirname(__file__),'template') # installed

# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily. 
NSPREFIXES = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships'
    }

FILES_TO_IGNORE = (
    'Thumbs.db', '.DS_STORE', 'document.xml', 'core.xml', 'app.xml', '[Content_Types].xml',
    'webSettings.xml', 'word/_rels/document.xml.rels',
)




def advReplace(document,search,replace,bs=3):
    '''Replace all occurences of string with a different string, return updated document
    
    This is a modified version of python-docx.replace() that takes into
    account blocks of <bs> elements at a time. The replace element can also
    be a string or an xml etree element.
    
    What it does:
    It searches the entire document body for text blocks.
    Then scan thos text blocks for replace.
    Since the text to search could be spawned across multiple text blocks,
    we need to adopt some sort of algorithm to handle this situation.
    The smaller matching group of blocks (up to bs) is then adopted.
    If the matching group has more than one block, blocks other than first
    are cleared and all the replacement text is put on first block.
    
    Examples:
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello,' / 'Hi!'
    output blocks : [ 'Hi!', '', ' world!' ]
    
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hello, world' / 'Hi!'
    output blocks : [ 'Hi!!', '', '' ]
    
    original text blocks : [ 'Hel', 'lo,', ' world!' ]
    search / replace: 'Hel' / 'Hal'
    output blocks : [ 'Hal', 'lo,', ' world!' ]
    
    @param instance  document: The original document
    @param str       search: The text to search for (regexp)
    @param mixed replace: The replacement text or lxml.etree element to
                          append, or a list of etree elements
    @param int       bs: See above
    
    @return instance The document with replacement applied
    
    '''
    # Enables debug output
    DEBUG = False
    
    newdocument = document
    
    # Compile the search regexp
    searchre = re.compile(search)
    
    # Will match against searchels. Searchels is a list that contains last
    # n text elements found in the document. 1 < n < bs
    searchels = []
    
    for element in newdocument.iter():
        if element.tag == '{%s}t' % NSPREFIXES['w']: # t (text) elements
            if element.text:
                # Add this element to searchels
                searchels.append(element)
                if len(searchels) > bs:
                    # Is searchels is too long, remove first elements
                    searchels.pop(0)
                
                # Search all combinations, of searchels, starting from
                # smaller up to bigger ones
                # l = search lenght
                # s = search start
                # e = element IDs to merge
                found = False
                for l in range(1,len(searchels)+1):
                    if found:
                        break
                    #print "slen:", l
                    for s in range(len(searchels)):
                        if found:
                            break
                        if s+l <= len(searchels):
                            e = range(s,s+l)
                            #print "elems:", e
                            txtsearch = ''
                            for k in e:
                                txtsearch += searchels[k].text
                
                            # Searcs for the text in the whole txtsearch
                            match = searchre.search(txtsearch)
                            if match:
                                found = True
                                
                                # I've found something :)
                                if DEBUG:
                                    print "Found element!"
                                    print "Search regexp:", searchre.pattern
                                    print "Requested replacement:", replace
                                    print "Matched text:", txtsearch
                                    print "Matched text (splitted):", map(lambda i:i.text,searchels)
                                    print "Matched at position:", match.start()
                                    print "matched in elements:", e
                                    if isinstance(replace, etree._Element):
                                        print "Will replace with XML CODE"
                                    elif type(replace) == list or type(replace) == tuple:
                                        print "Will replace with LIST OF ELEMENTS"
                                    else:
                                        print "Will replace with:", re.sub(search,replace,txtsearch)

                                curlen = 0
                                replaced = False
                                for i in e:
                                    curlen += len(searchels[i].text)
                                    if curlen > match.start() and not replaced:
                                        # The match occurred in THIS element. Puth in the
                                        # whole replaced text
                                        if isinstance(replace, etree._Element):
                                            # If I'm replacing with XML, clear the text in the
                                            # tag and append the element
                                            searchels[i].text = re.sub(search,'',txtsearch)
                                            searchels[i].append(replace)
                                        elif type(replace) == list or type(replace) == tuple:
                                            # I'm replacing with a list of etree elements
                                            searchels[i].text = re.sub(search,'',txtsearch)
                                            for r in replace:
                                                searchels[i].append(r)
                                        else:
                                            # Replacing with pure text
                                            searchels[i].text = re.sub(search,replace,txtsearch)
                                        replaced = True
                                        if DEBUG:
                                            print "Replacing in element #:", i
                                    else:
                                        # Clears the other text elements
                                        searchels[i].text = ''
    return newdocument
