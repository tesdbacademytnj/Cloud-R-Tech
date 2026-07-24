"""
Deep-dive: print raw hex around <a:blip> in header XML to find why regex fails.
"""
import re, zipfile
from io import BytesIO
from PIL import Image as PILImage
from lxml import etree
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def make_letterhead():
    img = PILImage.new('RGB', (200, 200), color='red')
    buf = BytesIO()
    img.save(buf, format='JPEG', quality=98)
    buf.seek(0)
    return buf

def add_full_page_background(doc, image_buf):
    NSMAP = {
        'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }
    W, WP, A, R, PIC = [NSMAP[k] for k in ('w', 'wp', 'a', 'r', 'pic')]

    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    rId, _ = header.part.get_or_add_image(image_buf)

    drawing = etree.Element(f'{{{W}}}drawing', nsmap=NSMAP)

    anchor = etree.SubElement(drawing, f'{{{WP}}}anchor')
    anchor.set('distT', '0')
    anchor.set('distB', '0')
    anchor.set('distL', '0')
    anchor.set('distR', '0')
    anchor.set('simplePos', '0')
    anchor.set('relativeHeight', '0')
    anchor.set('behindDoc', '1')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '1')

    simplePos = etree.SubElement(anchor, f'{{{WP}}}simplePos')
    simplePos.set('x', '0')
    simplePos.set('y', '0')

    pH = etree.SubElement(anchor, f'{{{WP}}}positionH')
    pH.set('relativeFrom', 'page')
    offH = etree.SubElement(pH, f'{{{WP}}}posOffset')
    offH.text = '0'

    pV = etree.SubElement(anchor, f'{{{WP}}}positionV')
    pV.set('relativeFrom', 'page')
    offV = etree.SubElement(pV, f'{{{WP}}}posOffset')
    offV.text = '0'

    extent = etree.SubElement(anchor, f'{{{WP}}}extent')
    extent.set('cx', str(int(595.56 * 12700)))
    extent.set('cy', str(int(842.52 * 12700)))

    etree.SubElement(anchor, f'{{{WP}}}wrapNone')

    docPr = etree.SubElement(anchor, f'{{{WP}}}docPr')
    docPr.set('id', '1')
    docPr.set('name', 'Background')

    etree.SubElement(anchor, f'{{{WP}}}cNvGraphicFramePr')

    graphic = etree.SubElement(anchor, f'{{{A}}}graphic')
    gData = etree.SubElement(graphic, f'{{{A}}}graphicData')
    gData.set('uri', PIC)

    pic = etree.SubElement(gData, f'{{{PIC}}}pic')
    nvPicPr = etree.SubElement(pic, f'{{{PIC}}}nvPicPr')
    cNvPr = etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPr')
    cNvPr.set('id', '1')
    cNvPr.set('name', 'Background')
    etree.SubElement(nvPicPr, f'{{{PIC}}}cNvPicPr')

    blipFill = etree.SubElement(pic, f'{{{PIC}}}blipFill')
    blip = etree.SubElement(blipFill, f'{{{A}}}blip')
    blip.set(f'{{{R}}}embed', rId)
    stretch = etree.SubElement(blipFill, f'{{{A}}}stretch')
    etree.SubElement(stretch, f'{{{A}}}fillRect')

    spPr = etree.SubElement(pic, f'{{{PIC}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{A}}}xfrm')
    off_el = etree.SubElement(xfrm, f'{{{A}}}off')
    off_el.set('x', '0')
    off_el.set('y', '0')
    ext_el = etree.SubElement(xfrm, f'{{{A}}}ext')
    ext_el.set('cx', str(int(595.56 * 12700)))
    ext_el.set('cy', str(int(842.52 * 12700)))
    prstGeom = etree.SubElement(spPr, f'{{{A}}}prstGeom')
    prstGeom.set('prst', 'rect')

    run = OxmlElement('w:r')
    run.append(drawing)

    bg_para = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '0')
    sp.set(qn('w:after'), '0')
    sp.set(qn('w:line'), '0')
    sp.set(qn('w:lineRule'), 'exact')
    pPr.append(sp)
    bg_para.append(pPr)
    bg_para.append(run)

    hdr_el = header._element
    existing = hdr_el.findall(qn('w:p'))
    if existing:
        hdr_el.insert(list(hdr_el).index(existing[0]), bg_para)
    else:
        hdr_el.append(bg_para)


doc = Document()
section = doc.sections[0]
section.page_width = Pt(595.56)
section.page_height = Pt(842.52)
p = doc.add_paragraph('Test content')

lh_buf = make_letterhead()
add_full_page_background(doc, lh_buf)

buf = BytesIO()
doc.save(buf)

with zipfile.ZipFile(BytesIO(buf.getvalue())) as z:
    for name in z.namelist():
        if name.endswith('.xml') and 'header' in name:
            raw = z.read(name)
            text = raw.decode('utf-8')

            print(f'=== {name} ===')
            print(f'Total length: {len(text)} chars')

            # Find the blip tag with surrounding context
            idx = text.find('a:blip')
            if idx >= 0:
                start = max(0, idx - 50)
                end = min(len(text), idx + 100)
                context = text[start:end]
                print(f'\nContext around "a:blip" (chars {idx-50}..{idx+100}):')
                print(repr(context))

                # Print each char as hex around the blip
                blip_area = text[idx-5:idx+60]
                print(f'\nChar-by-char around blip:')
                for i, ch in enumerate(blip_area):
                    print(f'  [{i}] ord={ord(ch):3d} hex={ord(ch):04x} char={repr(ch)}')

            # Now try the regex with detailed logging
            pattern = r'<a:blip (r:embed="rId\d+)/>'
            match = re.search(pattern, text)
            print(f'\nRegex search result: {match}')
            if match:
                print(f'  Matched: {match.group()}')
            else:
                # Try a simpler pattern
                simple = re.search(r'a:blip', text)
                if simple:
                    print(f'  Simple "a:blip" found at pos {simple.start()}')
                    print(f'  Surrounding: {repr(text[simple.start()-5:simple.start()+80])}')

            # Try the findall pattern
            all_blips = re.findall(r'<a:blip[^>]*>', text)
            print(f'\nfindall <a:blip[^>]*> results: {all_blips}')

            # Now try matching each blip findall result against the sub pattern
            for b in all_blips:
                sub_match = re.match(r'<a:blip (r:embed="rId\d+)/>', b)
                print(f'  Sub-match on {repr(b)}: {sub_match}')
