"""
Generate test DOCX files with different cstate values for comparison,
and check what Word actually does with them.
"""
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

# Generate the standard docx (with cstate="print")
docx_bytes = dg.generate_appraisal(test_data)

# Create variant WITHOUT cstate
def remove_cstate(data):
    buf_in = BytesIO(data)
    buf_out = BytesIO()
    with zipfile.ZipFile(buf_in, 'r') as zin, \
         zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            d = zin.read(name)
            if name.endswith('.xml'):
                text = d.decode('utf-8')
                text = text.replace(' cstate="print"', '')
                d = text.encode('utf-8')
            zout.writestr(name, d)
    buf_out.seek(0)
    return buf_out.getvalue()

# Create variant with cstate="screen"  
def change_cstate_to_screen(data):
    buf_in = BytesIO(data)
    buf_out = BytesIO()
    with zipfile.ZipFile(buf_in, 'r') as zin, \
         zipfile.ZipFile(buf_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            d = zin.read(name)
            if name.endswith('.xml'):
                text = d.decode('utf-8')
                text = text.replace('cstate="print"', 'cstate="screen"')
                d = text.encode('utf-8')
            zout.writestr(name, d)
    buf_out.seek(0)
    return buf_out.getvalue()

no_cstate = remove_cstate(docx_bytes)
screen_cstate = change_cstate_to_screen(docx_bytes)

base = os.path.join(os.path.dirname(__file__), '..')

with open(os.path.join(base, 'test_with_cstate_print.docx'), 'wb') as f:
    f.write(docx_bytes)
print(f"1. test_with_cstate_print.docx  ({len(docx_bytes)} bytes) - current version")

with open(os.path.join(base, 'test_without_cstate.docx'), 'wb') as f:
    f.write(no_cstate)
print(f"2. test_without_cstate.docx     ({len(no_cstate)} bytes) - no cstate attribute")

with open(os.path.join(base, 'test_with_cstate_screen.docx'), 'wb') as f:
    f.write(screen_cstate)
print(f"3. test_with_cstate_screen.docx ({len(screen_cstate)} bytes) - cstate=screen")

print("\nPlease open all three files in Microsoft Word and compare:")
print("  - Are colors visible in any of them?")
print("  - Which one shows the letterhead colors correctly?")
print("  - Which one shows the signature/seal colors correctly?")
