from lxml import etree
from docx import NSPREFIXES


def make_element(tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
    '''Create an element & return it''' 
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if type(nsprefix) == list:
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = NSPREFIXES[prefix]
        nsprefix = nsprefix[0] # FIXME: rest of code below expects a single prefix
    if nsprefix:
        namespace = '{'+NSPREFIXES[nsprefix]+'}'
    else:
        # For when namespace = None
        namespace = ''
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
            attributenamespace = '{'+NSPREFIXES[attrnsprefix]+'}'
                    
        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext    
    return newelement
