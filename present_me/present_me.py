
import pandas as pd
import os
import requests
from html2image import Html2Image
from pptx import Presentation
from pptx.util import Inches

folder_path = os.path.expanduser("~/Desktop/present_me")
excel_file = os.path.join(folder_path, "data.xlsx")
template_file = os.path.join(folder_path, "template.html")
output_folder = os.path.join(folder_path, "output_images")
resource_images_folder = os.path.join(folder_path, "resource_images")
os.makedirs(output_folder, exist_ok=True)
os.makedirs(resource_images_folder, exist_ok=True)

logo_left = os.path.join(folder_path, "logo_left.png")
logo_right = os.path.join(folder_path, "logo_right.png")
bottom_strip = os.path.join(folder_path, "bottom_strip.png")

with open(template_file, "r", encoding="utf-8") as f:
    template_html = f.read()

hti = Html2Image(output_path=output_folder)
hti.custom_flags = ["--window-size=2500,2000"]

def get_image_path(url, idx):
    if str(url).startswith("http://") or str(url).startswith("https://"):
        img_name = f"image_{idx+1}.jpg"
        img_path = os.path.join(resource_images_folder, img_name)
        try:
            r = requests.get(url, stream=True, timeout=10)
            if r.status_code == 200:
                with open(img_path, 'wb') as f:
                    for chunk in r:
                        f.write(chunk)
                return img_path
        except Exception as e:
            print(f"⚠️ Failed to download image {url}: {e}")
            return ""
    else:
        return os.path.join(folder_path, url)

def fix_local_images(html):
    local_images = { "{{logo_left}}": logo_left, "{{logo_right}}": logo_right, "{{bottom_strip}}": bottom_strip }
    for placeholder, path in local_images.items():
        html = html.replace(placeholder, path.replace("\\","/"))
    return html

def fill_template(template, row, idx):
    filled = template
    mapping = {
        "{{col2}}": row.get("col2",""),
        "{{col3}}": row.get("col3",""),
        "{{col4}}": row.get("col4",""),
        "{{col5}}": row.get("col5",""),
        "{{col6}}": row.get("col6",""),
        "{{col7}}": get_image_path(str(row.get("col7","")), idx)
    }
    for placeholder, value in mapping.items():
        filled = filled.replace(placeholder, str(value))
    return fix_local_images(filled)

def render_image(row, idx):
    html_content = fill_template(template_html, row, idx)
    output_path = os.path.join(output_folder, f"page_{idx+1}.png")
    hti.screenshot(html_str=html_content, save_as=os.path.basename(output_path))
    print(f"✅ Rendered page {idx+1} to {output_path}")
    return output_path

def add_blank_slide(prs):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    if os.path.exists(logo_left):
        slide.shapes.add_picture(logo_left, Inches(0.2), Inches(0.2), height=Inches(1))
    if os.path.exists(logo_right):
        slide.shapes.add_picture(logo_right, Inches(8), Inches(0.2), height=Inches(1))
    if os.path.exists(bottom_strip):
        slide.shapes.add_picture(bottom_strip, 0, prs.slide_height-Inches(1), width=prs.slide_width)
    return slide

def create_ppt(image_paths, ppt_file):
    prs = Presentation()
    add_blank_slide(prs)
    for img_path in image_paths:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
    add_blank_slide(prs)
    prs.save(ppt_file)
    print(f"✅ PPT saved: {ppt_file}")

print("=== Present Me PPT Generator ===")
sheet_names = pd.ExcelFile(excel_file).sheet_names
for sheet_name in sheet_names:
    print(f"Generating PPT for sheet: {sheet_name}")
    df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
    image_paths = [render_image(df_sheet.iloc[idx], idx) for idx in range(len(df_sheet))]
    ppt_file = os.path.join(folder_path, f"{sheet_name}.pptx")
    create_ppt(image_paths, ppt_file)
