"""Check letterhead image properties"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

from PIL import Image, ImageCms

img_path = os.path.join(os.path.dirname(__file__), '..', 'static', 'letters', 'img', 'letterhead.png')
if not os.path.exists(img_path):
    print(f"NOT FOUND: {img_path}")
    sys.exit(1)

img = Image.open(img_path)
print(f"File: {img_path}")
print(f"Size: {img.size}")
print(f"Mode: {img.mode}")
print(f"Format: {img.format}")
print(f"Info: {img.info}")

if 'icc_profile' in img.info:
    profile = ImageCms.ImageCmsProfile(img.info['icc_profile'])
    print(f"ICC Profile: {profile.profile.profile_description}")
else:
    print("ICC Profile: NONE")

# Check if it has alpha/transparency
if img.mode in ('RGBA', 'LA', 'PA'):
    print("HAS TRANSPARENCY - converting to JPEG will fill transparent areas with BLACK")
elif img.mode == 'P' and 'transparency' in img.info:
    print("HAS TRANSPARENCY (palette mode)")
else:
    print("No transparency")

# Also check pixel color at center
w, h = img.size
center = img.getpixel((w//2, h//2))
print(f"Center pixel RGBA: {center}")

# Check a few points
for label, pos in [('top-left', (10,10)), ('center', (w//2,h//2)), ('bottom-right', (w-10,h-10))]:
    px = img.getpixel(pos)
    print(f"  {label}: {px}")
