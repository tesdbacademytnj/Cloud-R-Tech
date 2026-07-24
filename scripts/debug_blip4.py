"""
Exact replication of _postprocess_docx_color with step-by-step debugging.
"""
import re, zipfile
from io import BytesIO
from PIL import Image as PILImage
from lxml import etree
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

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
    for k,v in [('distT','0'),('distB','0'),('distL','0'),('distR','0'),('simplePos','0'),('relativeHeight','0'),('behindDoc','1'),('locked','0'),('layoutInCell','1'),('allowOverlap','1')]:
        anchor.set(k, v)
    sp = etree.SubElement(anchor, f'{{{WP}}}simplePos'); sp.set('x','0'); sp.set('y','0')
    pH = etree.SubElement(anchor, f'{{{WP}}}positionH'); pH.set('relativeFrom','page')
    etree.SubElement(pH, f'{{{WP}}}posOffset').text = '0'
    pV = etree.SubElement(anchor, f'{{{WP}}}positionV'); pV.set('relativeFrom','page')
    etree.SubElement(pV, f'{{{WP}}}posOffset').text = '0'
    ext = etree.SubElement(anchor, f'{{{WP}}}extent'); ext.set('cx',str(int(595.56*12700))); ext.set('cy',str(int(842.52*12700)))
    etree.SubElement(anchor, f'{{{WP}}}wrapNone')
    dp = etree.SubElement(anchor, f'{{{WP}}}docPr'); dp.set('id','1'); dp.set('name','Background')
    etree.SubElement(anchor, f'{{{WP}}}cNvGraphicFramePr')
    gr = etree.SubElement(anchor, f'{{{A}}}graphic')
    gd = etree.SubElement(gr, f'{{{A}}}graphicData'); gd.set('uri', PIC)
    pic = etree.SubElement(gd, f'{{{PIC}}}pic')
    nvp = etree.SubElement(pic, f'{{{PIC}}}nvPicPr')
    cnvp = etree.SubElement(nvp, f'{{{PIC}}}cNvPr'); cnvp.set('id','1'); cnvp.set('name','Background')
    etree.SubElement(nvp, f'{{{PIC}}}cNvPicPr')
    bf = etree.SubElement(pic, f'{{{PIC}}}blipFill')
    blip = etree.SubElement(bf, f'{{{A}}}blip'); blip.set(f'{{{R}}}embed', rId)
    st = etree.SubElement(bf, f'{{{A}}}stretch'); etree.SubElement(st, f'{{{A}}}fillRect')
    sp2 = etree.SubElement(pic, f'{{{PIC}}}spPr')
    xf = etree.SubElement(sp2, f'{{{A}}}xfrm')
    oe = etree.SubElement(xf, f'{{{A}}}off'); oe.set('x','0'); oe.set('y','0')
    ee = etree.SubElement(xf, f'{{{A}}}ext'); ee.set('cx',str(int(595.56*12700))); ee.set('cy',str(int(842.52*12700)))
    pg = etree.SubElement(sp2, f'{{{A}}}prstGeom'); pg.set('prst','rect')

    run = OxmlElement('w:r'); run.append(drawing)
    bg = OxmlElement('w:p'); pp = OxmlElement('w:pPr')
    sp3 = OxmlElement('w:spacing')
    sp3.set(qn('w:before'),'0'); sp3.set(qn('w:after'),'0'); sp3.set(qn('w:line'),'0'); sp3.set(qn('w:lineRule'),'exact')
    pp.append(sp3); bg.append(pp); bg.append(run)
    hdr = header._element
    ex = hdr.findall(qn('w:p'))
    if ex:
        hdr.insert(list(hdr).index(ex[0]), bg)
    else:
        hdr.append(bg)


def _postprocess_docx_color(docx_bytes):
    """Exact copy from docx_generator.py"""
    import re as _re
    buf_in  = BytesIO(docx_bytes)
    buf_out = BytesIO()
    with zipfile.ZipFile(buf_in, 'r') as zin, \
         zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name.endswith('.xml'):
                content = data.decode('utf-8')
                old = content
                content = _re.sub(
                    r'<a:blip (r:embed="rId\d+")/>',
                    r'<a:blip \1 cstate="print"/>',
                    content
                )
                if content != old:
                    print(f"  [postprocess] MODIFIED: {name}")
                else:
                    print(f"  [postprocess] UNCHANGED: {name}")
                    # Debug: find all a:blip tags
                    blips = re.findall(r'<a:blip[^>]*>', content)
                    if blips:
                        print(f"    Found blip tags: {blips}")
                        for b in blips:
                            test = re.sub(r'<a:blip (r:embed="rId\d+")/>', r'<a:blip \1 cstate="print"/>', b)
                            print(f"    Direct sub on '{b}': '{test}' (changed: {test != b})")
                data = content.encode('utf-8')
            zout.writestr(name, data)
    buf_out.seek(0)
    return buf_out.getvalue()


# Build test doc
doc = Document()
section = doc.sections[0]
section.page_width = Pt(595.56)
section.page_height = Pt(842.52)
doc.add_paragraph('Test')

add_full_page_background(doc, make_letterhead())

buf = BytesIO()
doc.save(buf)
raw = buf.getvalue()

print("="*60)
print("  STEP 1: Raw docx has these blip tags:")
print("="*60)
with zipfile.ZipFile(BytesIO(raw)) as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            text = z.read(name).decode('utf-8')
            blips = re.findall(r'<a:blip[^>]*>', text)
            if blips:
                print(f"  {name}: {blips}")

print()
print("="*60)
print("  STEP 2: Running _postprocess_docx_color:")
print("="*60)
result = _postprocess_docx_color(raw)

print()
print("="*60)
print("  STEP 3: After post-processing:")
print("="*60)
with zipfile.ZipFile(BytesIO(result)) as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            text = z.read(name).decode('utf-8')
            blips = re.findall(r'<a:blip[^>]*>', text)
            cstate_count = text.count('cstate="print"')
            if blips or cstate_count:
                print(f"  {name}: blips={blips}, cstate_count={cstate_count}")
