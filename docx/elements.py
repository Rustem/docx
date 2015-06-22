import os

from PIL import Image
from lxml import etree

from utils import make_element


def pagebreak(type='page', orient='portrait'):
    '''Insert a break, default 'page'.
    See http://openxmldeveloper.org/forums/thread/4075.aspx
    Return our page break element.'''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']
    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))
    pagebreak = make_element('p')
    if type == 'page':
        run = make_element('r')
        br = make_element('br',attributes={'type':type})
        run.append(br)
        pagebreak.append(run)
    elif type == 'section':
        pPr = make_element('pPr')
        sectPr = make_element('sectPr')
        if orient == 'portrait':
            pgSz = make_element('pgSz',attributes={'w':'12240','h':'15840'})
        elif orient == 'landscape':
            pgSz = make_element('pgSz',attributes={'h':'12240','w':'15840', 'orient':'landscape'})
        sectPr.append(pgSz)
        pPr.append(sectPr)
        pagebreak.append(pPr)
    return pagebreak


def paragraph(paratext,style='BodyText',breakbefore=False,jc='left'):
    '''Make a new paragraph element, containing a run, and some text. 
    Return the paragraph element.
    
    @param string jc: Paragraph alignment, possible values:
                      left, center, right, both (justified), ...
                      see http://www.schemacentral.com/sc/ooxml/t-w_ST_Jc.html
                      for a full list
    
    If paratext is a list, spawn multiple run/text elements.
    Support text styles (paratext must then be a list of lists in the form
    <text> / <style>. Stile is a string containing a combination od 'bui' chars
    
    example
    paratext = [
        ['some bold text', 'b'],
        ['some normal text', ''],
        ['some italic underlined text', 'iu'],
    ]
    
    '''
    # Make our elements
    paragraph = make_element('p')
    
    if type(paratext) == list:
        text = []
        for pt in paratext:
            if type(pt) == list:
                text.append([make_element('t',tagtext=pt[0]), pt[1]])
            else:
                text.append([make_element('t',tagtext=pt), ''])
    else:
        text = [[make_element('t',tagtext=paratext),''],]
    pPr = make_element('pPr')
    pStyle = make_element('pStyle',attributes={'val':style})
    pJc = make_element('jc',attributes={'val':jc})
    pPr.append(pStyle)
    pPr.append(pJc)
                
    # Add the text the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = make_element('r')
        rPr = make_element('rPr')
        # Apply styles
        if t[1].find('b') > -1:
            b = make_element('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = make_element('u',attributes={'val':'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = make_element('i')
            rPr.append(i)
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
            lastRenderedPageBreak = make_element('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph


def heading(headingtext,headinglevel,lang='en'):
    '''Make a new heading, return the heading element'''
    lmap = {
        'en': 'Heading',
        'it': 'Titolo',
    }
    # Make our elements
    paragraph = make_element('p')
    pr = make_element('pPr')
    pStyle = make_element('pStyle',attributes={'val':lmap[lang]+str(headinglevel)})    
    run = make_element('r')
    text = make_element('t',tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)   
    paragraph.append(run)    
    # Return the combined paragraph
    return paragraph


def table(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
    '''Get a list of lists, return a table
    
        @param list contents: A list of lists describing contents
                              Every item in the list can be a string or a valid
                              XML element itself. It can also be a list. In that case
                              all the listed elements will be merged into the cell.
        @param bool heading: Tells whether first line should be threated as heading
                             or not
        @param list colw: A list of interger. The list must have same element
                          count of content lines. Specify column Widths in
                          wunitS
        @param string cwunit: Unit user for column width:
                                'pct': fifties of a percent
                                'dxa': twenties of a point
                                'nil': no width
                                'auto': automagically determined
        @param int tblw: Table width
        @param int twunit: Unit used for table width. Same as cwunit
        @param dict borders: Dictionary defining table border. Supported keys are:
                             'top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'all'
                             When specified, the 'all' key has precedence over others.
                             Each key must define a dict of border attributes:
                             color: The color of the border, in hex or 'auto'
                             space: The space, measured in points
                             sz: The size of the border, in eights of a point
                             val: The style of the border, see http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
        @param list celstyle: Specify the style for each colum, list of dicts.
                              supported keys:
                              'align': specify the alignment, see paragraph documentation,
        
        @return lxml.etree: Generated XML etree element
    '''
    table = make_element('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = make_element('tblPr')
    tablestyle = make_element('tblStyle',attributes={'val':'ColorfulGrid-Accent1'})
    tableprops.append(tablestyle)
    tablewidth = make_element('tblW',attributes={'w':str(tblw),'type':str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = make_element('tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
                attrs = {}
                for a in borders[k].keys():
                    attrs[a] = unicode(borders[k][a])
                borderelem = make_element(b,attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = make_element('tblLook',attributes={'val':'0400'})
    tableprops.append(tablelook)
    table.append(tableprops)    
    # Table Grid    
    tablegrid = make_element('tblGrid')
    for i in range(columns):
        tablegrid.append(make_element('gridCol',attributes={'w':str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)     
    # Heading Row    
    row = make_element('tr')
    rowprops = make_element('trPr')
    cnfStyle = make_element('cnfStyle',attributes={'val':'000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    if heading:
        i = 0
        for heading in contents[0]:
            cell = make_element('tc')  
            # Cell properties  
            cellprops = make_element('tcPr')
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
            cellwidth = make_element('tcW',attributes=wattr)
            cellstyle = make_element('shd',attributes={'val':'clear','color':'auto','fill':'548DD4','themeFill':'text2','themeFillTint':'99'})
            cellprops.append(cellwidth)
            cellprops.append(cellstyle)
            cell.append(cellprops)        
            # Paragraph (Content)
            if not type(heading) == list and not type(heading) == tuple:
                heading = [heading,]
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
                    cell.append(paragraph(h,jc='center'))
            row.append(cell)
            i += 1
        table.append(row)          
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = make_element('tr')     
        i = 0
        for content in contentrow:   
            cell = make_element('tc')
            # Properties
            cellprops = make_element('tcPr')
            if colw:
                wattr = {'w':str(colw[i]),'type':cwunit}
            else:
                wattr = {'w':'0','type':'auto'}
            cellwidth = make_element('tcW',attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not type(content) == list and not type(content) == tuple:
                content = [content,]
            for c in content:
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    if celstyle and 'align' in celstyle[i].keys():
                        align = celstyle[i]['align']
                    else:
                        align = 'left'
                    cell.append(paragraph(c,jc=align))
            row.append(cell)    
            i += 1
        table.append(row)   
    return table


def picture(document, picname, picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
    '''Take a relationshiplist, picture file name, and return a paragraph containing the image
    and an updated relationshiplist'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''  
    document.word_relationships.to_copy.append([picname, os.path.abspath(picname)])

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth,pixelheight = Image.open(picname).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs            
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)   
    
    # Set relationship ID to the first available  
    picid = '2'
    picrelid = 'rId'+str(len(document.word_relationships.relationshiplist) + 1)
    picid = str(len(document.word_relationships.relationshiplist) + 1)
    document.word_relationships.relationshiplist.append([
        picrelid, 
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])
    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = make_element('blipFill',nsprefix='pic')
    blipfill.append(make_element('blip',nsprefix='a',attrnsprefix='r',attributes={'embed':picrelid}))
    stretch = make_element('stretch',nsprefix='a')
    stretch.append(make_element('fillRect',nsprefix='a'))
    blipfill.append(make_element('srcRect',nsprefix='a'))
    blipfill.append(stretch)
    
    # 2. The non visual picture properties 
    nvpicpr = make_element('nvPicPr',nsprefix='pic')
    cnvpr = make_element('cNvPr',nsprefix='pic',
                        attributes={'id':'0','name':'Picture 1','descr':picname}) 
    nvpicpr.append(cnvpr) 
    cnvpicpr = make_element('cNvPicPr',nsprefix='pic')                           
    cnvpicpr.append(make_element('picLocks', nsprefix='a', 
                    attributes={'noChangeAspect':str(int(nochangeaspect)),
                    'noChangeArrowheads':str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)
        
    # 3. The Shape properties
    sppr = make_element('spPr',nsprefix='pic',attributes={'bwMode':'auto'})
    xfrm = make_element('xfrm',nsprefix='a')
    xfrm.append(make_element('off',nsprefix='a',attributes={'x':'0','y':'0'}))
    xfrm.append(make_element('ext',nsprefix='a',attributes={'cx':width,'cy':height}))
    prstgeom = make_element('prstGeom',nsprefix='a',attributes={'prst':'rect'})
    prstgeom.append(make_element('avLst',nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)
    
    # Add our 3 parts to the picture element
    pic = make_element('pic',nsprefix='pic')    
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)
    
    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = make_element('graphicData',nsprefix='a',
        attributes={'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = make_element('graphic',nsprefix='a')
    graphic.append(graphicdata)

    framelocks = make_element('graphicFrameLocks',nsprefix='a',attributes={'noChangeAspect':'1'})    
    framepr = make_element('cNvGraphicFramePr',nsprefix='wp')
    framepr.append(framelocks)
    docpr = make_element('docPr',nsprefix='wp',
        attributes={'id':picid,'name':'Picture 1','descr':picdescription})
    effectextent = make_element('effectExtent',nsprefix='wp',
        attributes={'l':'25400','t':'0','r':'0','b':'0'})
    extent = make_element('extent',nsprefix='wp',attributes={'cx':width,'cy':height})
    inline = make_element('inline',
        attributes={'distT':"0",'distB':"0",'distL':"0",'distR':"0"},nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = make_element('drawing')
    drawing.append(inline)
    run = make_element('r')
    run.append(drawing)
    paragraph = make_element('p')
    paragraph.append(run)
    return paragraph



