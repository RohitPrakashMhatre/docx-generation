from docx import Document
from docx.shared import Inches
import requests
from PIL import Image, ImageOps
from io import BytesIO  # <- correct name (not BytestIO)
# 
# --- resize helper (high quality, no upscaling) ---
def resize_image(img: Image.Image, max_width=1200, max_height=1200) -> Image.Image:
    # fix EXIF rotation if present
    img = ImageOps.exif_transpose(img)

    w, h = img.size
    if w <= max_width and h <= max_height:
        return img  # no change

    ratio = min(max_width / float(w), max_height / float(h))
    new_size = (int(w * ratio), int(h * ratio))  # <- was height*height (bug)
    return img.resize((new_size), Image.LANCZOS)

# --- replace placeholder with an image (file path or BytesIO) ---
def replace_placeholder_with_image(paragraph, placeholder, image_stream_or_path, width_inches=5.0):
    # 1) combine runs to find placeholder even if split across runs
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder not in full_text:
        return False

    # 2) split at first occurrence
    before, _, after = full_text.partition(placeholder)

    # 3) clear runs, keep first for 'before'
    if paragraph.runs:
        paragraph.runs[0].text = before
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.add_run(before)

    # 4) insert image
    run = paragraph.add_run()
    run.add_picture(image_stream_or_path, width=Inches(width_inches))

    # 5) append trailing text
    paragraph.add_run(after)
    return True

# ----------------- main -----------------
match = "{{floor_img.plan}}"
image_url = "https://app-pythondoc.bluealgo.com/file_storage_downloads/20251018102533880643__Miraya.jpg"
doc = Document("main_template.docx")

# download image
resp = requests.get(image_url, timeout=30)
resp.raise_for_status()

# open and resize (in-memory)
img = Image.open(BytesIO(resp.content))
img = resize_image(img, max_width=1200, max_height=1200)
img.save("resized_preview.jpg", format="JPEG", quality=95, subsampling=0, optimize=True)

print("Image size: ", img.size)

# save to an in-memory stream that python-docx can read
img_buf = BytesIO()
# JPEG for photos; PNG if you need transparency (Word inlines PNG fine)
if img.mode in ("RGBA", "LA"):
    img = img.convert("RGB")
img.save(img_buf, format="png", quality=95, subsampling=0, optimize=True)
img_buf.seek(0)

# replace in paragraphs
for p in doc.paragraphs:
    if match in p.text:
        replaced = replace_placeholder_with_image(p, match, img_buf, width_inches=5.0)
        # if only first occurrence matters per paragraph, you can skip further checks

doc.save("output.docx")
print("✅ Word file generated successfully: output.docx")
