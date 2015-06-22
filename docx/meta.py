import time
from lxml import etree
from docx.utils import make_element

class CoreProperties(object):
    """Core properties for a document.
    """
    def __init__(self, title, creator, subject='', keywords=[], lastmodifiedby=None):
        self.title = title
        self.creator = creator
        self.subject = subject
        self.keywords = keywords
        self.lastmodifiedby = lastmodifiedby if lastmodifiedby is not None else creator

    def _xml(self):
        coreprops = make_element('coreProperties',nsprefix='cp')    
        coreprops.append(make_element('title', tagtext=self.title, nsprefix='dc'))
        coreprops.append(make_element('subject', tagtext=self.subject, nsprefix='dc'))
        coreprops.append(make_element('creator', tagtext=self.creator, nsprefix='dc'))
        coreprops.append(make_element('keywords', tagtext=','.join(self.keywords), nsprefix='cp'))    
        coreprops.append(make_element('lastModifiedBy', tagtext=self.lastmodifiedby, nsprefix='cp'))
        coreprops.append(make_element('revision', tagtext='1', nsprefix='cp'))
        coreprops.append(make_element('category', tagtext='Examples', nsprefix='cp'))
        coreprops.append(make_element('description', tagtext='Examples', nsprefix='dc'))
        currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')
        # Document creation and modify times
        # Prob here: we have an attribute who name uses one namespace, and that 
        # attribute's value uses another namespace.
        # We're creating the element from a string as a workaround...
        for doctime in ['created','modified']:
            coreprops.append(etree.fromstring('''<dcterms:'''+doctime+''' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcterms="http://purl.org/dc/terms/" xsi:type="dcterms:W3CDTF">'''+currenttime+'''</dcterms:'''+doctime+'''>'''))
            pass
        return coreprops


class AppProperties(object):
    """Properties describing the application which created the OpenXML file."""

    def __init__(self, application='Microsoft Word 12.0.0', version='12.000'):
        self.application=application
        self.version=version

    def _xml(self):
        appprops = make_element('Properties',nsprefix='ep')
        appprops = etree.fromstring(
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"></Properties>''')
        props = {
                'Template':'Normal.dotm',
                'TotalTime':'6',
                'Pages':'1',  
                'Words':'83',   
                'Characters':'475', 
                'Application':self.application,
                'DocSecurity':'0',
                'Lines':'12', 
                'Paragraphs':'8',
                'ScaleCrop':'false', 
                'LinksUpToDate':'false', 
                'CharactersWithSpaces':'583',  
                'SharedDoc':'false',
                'HyperlinksChanged':'false',
                'AppVersion':self.version,
                }
        for prop in props:
            appprops.append(make_element(prop,tagtext=props[prop],nsprefix=None))
        return appprops


class WordRelationships(object):

    def __init__(self, xml=None):
        # Keeping track of which files have been added and copying them into the zipfile on saving.
        # Not a great solution but want to avoid temp_dirs and/or copying the file into the template dir.
        self.to_copy = []
        if xml:
            tree = etree.fromstring(xml)
            self.relationshiplist = [[r.get('Id'), r.get('Type'), r.get('Target')] for r in list(tree)]
        else:
            self.relationshiplist = [
            ['rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering','numbering.xml'],
            ['rId2', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles','styles.xml'],
            ['rId3', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml'],
            ['rId4', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings','webSettings.xml'],
            ['rId5', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable','fontTable.xml'],
            ['rId6', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme','theme/theme1.xml'],
        ]

    def _xml(self):
        '''Generate a Word relationships file'''
        # FIXME: using string hack instead of making element
        relationships = etree.fromstring(
        '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            </Relationships>'''
        )
        for relationship in self.relationshiplist:
            relationships.append(make_element('Relationship',attributes={'Id': relationship[0],
            'Type':relationship[1],'Target':relationship[2]},nsprefix=None))
        return relationships


class ContentTypes(object):
    # FIXME - doesn't quite work...read from string as temp hack...
    #types = make_element('Types',nsprefix='ct')
    def __init__(self, xml=None):
        if xml:
            self.types = dict()
            tree = etree.fromstring(xml)
            for r in list(tree):
                if 'Override' in r.tag:
                    self.types[r.get('PartName')] = r.get('ContentType')
        else:
            self.types = {
                '/word/theme/theme1.xml':'application/vnd.openxmlformats-officedocument.theme+xml',
                '/word/fontTable.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
                '/docProps/core.xml':'application/vnd.openxmlformats-package.core-properties+xml',
                '/docProps/app.xml':'application/vnd.openxmlformats-officedocument.extended-properties+xml',
                '/word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
                '/word/settings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
                '/word/numbering.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
                '/word/styles.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
                '/word/webSettings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'
                }

    def _xml(self):
        content_types = etree.fromstring('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>''')
        for t in self.types:
            content_types.append(
                make_element('Override',nsprefix=None,attributes={'PartName':t,'ContentType':self.types[t]})
            )
        # Add support for filetypes
        filetypes = {
            'rels':'application/vnd.openxmlformats-package.relationships+xml',
            'xml':'application/xml',
            'jpeg':'image/jpeg',
            'gif':'image/gif',
            'png':'image/png',
            'wmf': 'image/x-wmf',
        }
        for extension in filetypes:
            content_types.append(
                make_element(
                    'Default',
                    nsprefix=None,
                    attributes={
                        'Extension':extension,
                        'ContentType':filetypes[extension]
                    }
                )
            )
        return content_types


class WebSettings():
    '''Generate websettings'''
    def __init__(self):
        pass

    def _xml(self):
        web = make_element('webSettings')
        web.append(make_element('allowPNG'))
        web.append(make_element('doNotSaveAsSingleFile'))
        return web
