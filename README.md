# Present Me - A PPT Generator

**Present Me** is a Python-based command-line tool to generate PPT presentations automatically from an Excel workbook and an HTML template. It is designed to be **industry-agnostic** and can be used for marketing, sales, education, or any kind of presentations.


## Features

- Generate PPT slides from Excel rows automatically.  
- Render HTML content as images for slides using **html2image**.  
- Support multiple PPTs from a single Excel workbook (one sheet per PPT).  
- Include logos, bottom strip, and custom images.  
- Optional interactive slide replacement after PPT creation.  


## Getting Started

### 1. Download the required files

- Download `data.xlsx` (Excel workbook) exactly as it is.  
- Download `template.html` (HTML template) exactly as it is.  
- Place them in the same folder as the Python script (`present_me.py`).
- In the terminal, navigate to cd ~/desktop/present_me or wherever your folder is.
- Open a virtual env and install the packages. (pip install...)

✅ Add the required local images

This script uses three optional branding assets:
- logo_left.png
- logo_right.png
- bottom_strip.png

### Where these files must be placed
All three image files must be placed inside the same folder defined as:
folder_path = os.path.expanduser("~/Desktop/artwork")

**Important:**  

- The Excel columns start at `col2`. You can add a `col1` in your workbook for your own reference, but it will **not affect the PPT generation**.  
- Use a **separate sheet in the workbook for each PPT** you want to generate. The script will create one PPT per sheet.  

### 2. Set up a Python virtual environment

It is recommended to run Present Me in a virtual environment to avoid conflicts with other packages.  

### 3. Install the requirements

pip install --upgrade pip
pip install pandas requests html2image python-pptx openpyxl

### 4. Prepare your Excel workbook

- Each row represents one slide in the PPT.  

- **Important:** The Excel columns start at `col2`. You **can add a `col1`** for your own reference, numbering, or notes — it **will not affect the PPT generation**.  

- Columns used by the script:

| Column | Purpose |
|--------|---------|
| col2   | Page number |
| col3   | Slide title |
| col4   | Subtitle / Presenter / Author |
| col5   | Description / Features |
| col6   | Notes / Call to Action |
| col7   | Image URL or local path |

- Each sheet in the workbook corresponds to **one PPT**.


### 5. Run the script!

Use:

python present_me.py


The script will:
-Read each sheet in the Excel workbook.
-Render HTML slides as images.
-Create a PPT with all slides and add logos and bottom strip.
-Optionally allow interactive slide replacement.
-Rendered images are saved in the output_images folder.
-Generated PPT files are saved in the same folder as the script, named after the sheet names.

