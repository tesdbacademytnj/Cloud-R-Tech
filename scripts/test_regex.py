import re

# The actual XML text from the header
test_xml = '<pic:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'

# Current regex from docx_generator.py
pattern = r'<a:blip (r:embed="rId\d+")/>'
replacement = r'<a:blip \1 cstate="print"/>'

print("Input XML:")
print(test_xml)
print()

# Test regex
m = re.search(pattern, test_xml)
print(f"re.search result: {m}")
if m:
    print(f"  Matched: {m.group()}")

result = re.sub(pattern, replacement, test_xml)
print(f"\nAfter re.sub:")
print(result)
print(f"Changed: {result != test_xml}")

# Test with alternative patterns
print("\n--- Alt pattern 1: match any a:blip tag ---")
p2 = r'(<a:blip\b[^>]*?)(/>)'
m2 = re.search(p2, test_xml)
print(f"Result: {m2}")
if m2:
    print(f"  Matched: {m2.group()}")

# Test with the exact sub
result2 = re.sub(r'(<a:blip\b[^>]*?)(/>)', r'\1 cstate="print"\2', test_xml)
print(f"After alt sub: {result2}")
print(f"Changed: {result2 != test_xml}")
