"""Test the ACTUAL docx_generator module, not a replica."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

import django
django.setup()

from io import BytesIO
import re, zipfile
from letters import docx_generator as dg

# Generate a test appraisal docx
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

print("Generating appraisal DOCX...")
docx_bytes = dg.generate_appraisal(test_data)
print(f"Generated: {len(docx_bytes)} bytes")

# Inspect the zip
with zipfile.ZipFile(BytesIO(docx_bytes)) as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            text = z.read(name).decode('utf-8')
            blips = re.findall(r'<a:blip[^>]*>', text)
            cstate_count = text.count('cstate="print"')
            if blips or cstate_count:
                print(f"\n  {name}:")
                print(f"    blip tags: {len(blips)}, cstate count: {cstate_count}")
                for b in blips:
                    has_cstate = 'cstate' in b
                    print(f"    {b}  <-- cstate: {has_cstate}")

# Save for manual inspection
out_path = os.path.join(os.path.dirname(__file__), '..', 'test_output_appraisal.docx')
with open(out_path, 'wb') as f:
    f.write(docx_bytes)
print(f"\nSaved to: {out_path}")
