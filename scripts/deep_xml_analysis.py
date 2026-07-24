"""Deep XML analysis of the actual generated DOCX - full header and document XML"""
import os, sys, re, zipfile
from io import BytesIO
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
import django
django.setup()

from letters import docx_generator as dg

test_data = {
    'name': 'Test User',
    'gender': 'male',
    'address': '123 Test Street',
    'pincode': '600001',
    'date': '01 Jan 2025',
    'current_monthly': 50000,
    'new_monthly': 60000,
    'designation': 'Software Engineer',
    'ctc_lpa': '8',
    'hr_name': 'Raj Padmanaban',
    'sig_src': None,
}

docx_bytes = dg.generate_appraisal(test_data)

with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            text = z.read(name).decode('utf-8')
            if 'blip' in text.lower() or 'anchor' in text.lower() or 'drawing' in text.lower():
                print(f"\n{'='*70}")
                print(f"FILE: {name}")
                print(f"{'='*70}")
                
                # Pretty-print the relevant sections
                # Find all drawing/anchor sections
                anchors = re.findall(r'<wp:anchor[^>]*>.*?</wp:anchor>', text, re.DOTALL)
                for i, a in enumerate(anchors):
                    print(f"\n--- Anchor {i+1} ---")
                    # Extract key attributes
                    relH = re.search(r'relativeHeight="(\d+)"', a)
                    behindDoc = re.search(r'behindDoc="(\d+)"', a)
                    blip_embed = re.search(r'<a:blip[^>]*r:embed="(rId\d+)"[^>]*/>', a)
                    cstate = re.search(r'cstate="(\w+)"', a)
                    
                    print(f"  relativeHeight: {relH.group(1) if relH else 'NOT FOUND'}")
                    print(f"  behindDoc: {behindDoc.group(1) if behindDoc else 'NOT FOUND'}")
                    print(f"  blip r:embed: {blip_embed.group(1) if blip_embed else 'NOT FOUND'}")
                    print(f"  cstate: {cstate.group(1) if cstate else 'NOT FOUND'}")
                    
                    # Check positionH/V
                    posH = re.search(r'<wp:positionH[^>]*relativeFrom="([^"]*)"', a)
                    posV = re.search(r'<wp:positionV[^>]*relativeFrom="([^"]*)"', a)
                    print(f"  positionH relativeFrom: {posH.group(1) if posH else 'NOT FOUND'}")
                    print(f"  positionV relativeFrom: {posV.group(1) if posV else 'NOT FOUND'}")
                    
                    # Check extent
                    ext_cx = re.search(r'<wp:extent[^>]*cx="(\d+)"', a)
                    ext_cy = re.search(r'<wp:extent[^>]*cy="(\d+)"', a)
                    if ext_cx and ext_cy:
                        print(f"  extent cx={ext_cx.group(1)} cy={ext_cy.group(1)}")
                    
                    # Check wrapNone
                    has_wrapNone = 'wrapNone' in a
                    print(f"  wrapNone: {has_wrapNone}")
                    
                    # Print a compact version of the anchor
                    compact = re.sub(r'\s+', ' ', a)
                    if len(compact) > 500:
                        print(f"  XML (compact, truncated): {compact[:500]}...")
                    else:
                        print(f"  XML (compact): {compact}")
                
                # Check for inline drawings too
                inlines = re.findall(r'<wp:inline[^>]*>.*?</wp:inline>', text, re.DOTALL)
                for i, il in enumerate(inlines):
                    print(f"\n--- Inline {i+1} ---")
                    compact = re.sub(r'\s+', ' ', il)
                    if len(compact) > 300:
                        print(f"  XML (compact, truncated): {compact[:300]}...")
                    else:
                        print(f"  XML (compact): {compact}")
                    
                    # Check blip
                    blip_embed = re.search(r'r:embed="(rId\d+)"', il)
                    cstate = re.search(r'cstate="(\w+)"', il)
                    print(f"  blip r:embed: {blip_embed.group(1) if blip_embed else 'NOT FOUND'}")
                    print(f"  cstate: {cstate.group(1) if cstate else 'NOT FOUND'}")
