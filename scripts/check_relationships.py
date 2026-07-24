"""Check image relationships and [Content_Types].xml in the generated DOCX"""
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
    # Check relationships for header
    print("=== word/_rels/header1.xml.rels ===")
    try:
        rels = z.read('word/_rels/header1.xml.rels').decode('utf-8')
        print(rels)
    except KeyError:
        print("NOT FOUND")

    print("\n=== word/_rels/document.xml.rels ===")
    try:
        rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
        print(rels)
    except KeyError:
        print("NOT FOUND")

    print("\n=== [Content_Types].xml ===")
    ct = z.read('[Content_Types].xml').decode('utf-8')
    # Pretty print
    ct = ct.replace('><', '>\n<')
    for line in ct.split('\n'):
        if 'image' in line.lower() or 'jpeg' in line.lower() or 'png' in line.lower():
            print(f"  {line.strip()}")

    # Check image files
    print("\n=== Image files in zip ===")
    for name in z.namelist():
        if name.startswith('word/media/'):
            data = z.read(name)
            print(f"  {name}: {len(data)} bytes")
            # Check magic bytes
            if data[:2] == b'\xff\xd8':
                print(f"    Format: JPEG")
                # Check for ICC profile
                if b'ICC_PROFILE' in data:
                    print(f"    Has ICC profile: YES")
                else:
                    print(f"    Has ICC profile: NO (checking by offset)")
                    # Search for ICC profile marker
                    idx = data.find(b'ICC_PROFILE')
                    if idx >= 0:
                        print(f"    ICC_PROFILE found at offset {idx}")
                    else:
                        print(f"    No ICC_PROFILE marker found in JPEG data")
            elif data[:8] == b'\x89PNG\r\n\x1a\n':
                print(f"    Format: PNG")
                if b'cHRM' in data:
                    print(f"    Has cHRM chunk: YES")
                if b'iCCP' in data:
                    print(f"    Has iCCP chunk: YES")
                if b'sRGB' in data:
                    print(f"    Has sRGB chunk: YES")
            else:
                print(f"    Format: UNKNOWN (magic: {data[:8].hex()})")

    # Check actual blip elements and their image data
    print("\n=== Blip-to-image mapping ===")
    for xml_file in ['word/header1.xml', 'word/document.xml']:
        try:
            text = z.read(xml_file).decode('utf-8')
        except KeyError:
            continue
        blips = re.findall(r'<a:blip[^>]*r:embed="(rId\d+)"[^>]*/>', text)
        print(f"\n  {xml_file}:")
        
        # Parse relationships
        rels_file = xml_file.replace('.xml', '.xml.rels').replace('word/', 'word/_rels/')
        try:
            rels_text = z.read(rels_file).decode('utf-8')
        except KeyError:
            print(f"    (no rels file)")
            continue
            
        for rid in blips:
            # Find target in rels
            pattern = rf'Id="{rid}"[^>]*Target="([^"]*)"'
            match = re.search(pattern, rels_text)
            if match:
                target = match.group(1)
                target_path = f"word/{target}"
                print(f"    {rid} -> {target}")
                if target_path in z.namelist():
                    img_data = z.read(target_path)
                    fmt = "JPEG" if img_data[:2] == b'\xff\xd8' else "PNG" if img_data[:4] == b'\x89PNG' else "UNKNOWN"
                    print(f"      File: {target_path} ({len(img_data)} bytes, {fmt})")
                    
                    # Check if the JPEG has ICC profile
                    if fmt == "JPEG":
                        has_icc = b'ICC_PROFILE' in img_data
                        print(f"      ICC_PROFILE in data: {has_icc}")
            else:
                print(f"    {rid} -> (not found in rels)")
