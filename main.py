# Cell 1: Environment Setup
!apt-get -qq update --no-install-recommends
!apt-get -qq install -y exiftool libimage-exiftool-perl
!pip install -q -U "gradio" "openpyxl" "pillow" "pandas<3.0.0
# Cell 2: Metadata Extractor with ZIP Support & Embedded Thumbnails
import gradio as gr
import os
import re
import shutil
import tempfile
import zipfile
import json
import subprocess
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def extract_all_exif(path):
    try:
        cmd = [
            "exiftool",
            "-j",
            "-s",
            "-a",
            "-ee",
            "-c", "%.6f",   # Formats GPS coordinates to clean decimal degrees
            path
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            data = json.loads(result.stdout)
            return data[0] if data else {}
    except Exception as e:
        print(f"Error reading {path}: {e}")
    return {}


def is_valid_val(val):
    if val is None:
        return False
    s = str(val).strip().lower()
    return s not in ["", "nan", "unknown", "unknown ()", "undef", "0", "0.0"]


def parse_photo_details(source_path, raw_meta):
    filename = os.path.basename(source_path)
    name_without_ext, ext = os.path.splitext(filename)

    # 1. Device Make & Model
    make = raw_meta.get("Make") or raw_meta.get("EXIF:Make") or ""
    model = raw_meta.get("Model") or raw_meta.get("EXIF:Model") or raw_meta.get("XiaomiModel") or ""
    device = f"{make} {model}".strip() or "N/A"

    # 2. Date & Time Parsing
    dt_orig = (
        raw_meta.get("DateTimeOriginal")
        or raw_meta.get("CreateDate")
        or raw_meta.get("ModifyDate")
        or raw_meta.get("SubSecDateTimeOriginal")
    )
    date_val, time_val = "N/A", "N/A"
    if dt_orig and isinstance(dt_orig, str) and " " in dt_orig:
        parts = dt_orig.split(" ", 1)
        date_val = parts[0].replace(":", "-")
        time_val = parts[1].split("+")[0].strip()
    else:
        # Fallback to Android filename timestamp
        match = re.search(r"(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})", filename)
        if match:
            date_val = f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
            time_val = f"{match.group(4)}:{match.group(5)}:{match.group(6)}"

    # 3. GPS Coordinates
    gps_val = raw_meta.get("GPSPosition") or raw_meta.get("Composite:GPSPosition")
    if not is_valid_val(gps_val):
        lat = raw_meta.get("GPSLatitude")
        lon = raw_meta.get("GPSLongitude")
        lat_ref = raw_meta.get("GPSLatitudeRef", "")
        lon_ref = raw_meta.get("GPSLongitudeRef", "")

        lat_ref = "" if "unknown" in str(lat_ref).lower() else str(lat_ref).strip()
        lon_ref = "" if "unknown" in str(lon_ref).lower() else str(lon_ref).strip()

        if is_valid_val(lat) and is_valid_val(lon):
            gps_val = f"{lat} {lat_ref}, {lon} {lon_ref}".strip()
        else:
            gps_val = "Not Detected"

    # 4. Focal Length, Aperture, Brightness, Shutter & APEX Values
    focal_length = raw_meta.get("FocalLength35efl") or raw_meta.get("FocalLength") or "N/A"
    f_number = raw_meta.get("FNumber") or raw_meta.get("Aperture") or raw_meta.get("ApertureValue") or "N/A"
    brightness = raw_meta.get("BrightnessValue") or raw_meta.get("LightValue") or "N/A"
    apex_shutter = raw_meta.get("ShutterSpeed") or raw_meta.get("ExposureTime") or raw_meta.get("ShutterSpeedValue") or "N/A"
    apex_aperture = raw_meta.get("ApertureValue") or raw_meta.get("FNumber") or "N/A"
    iso = raw_meta.get("ISO") or "N/A"
    exp_bias = raw_meta.get("ExposureCompensation") or raw_meta.get("ExposureBiasValue") or "N/A"

    standard_data = {
        "Thumbnail": "",
        "File Name": filename,
        "Extension": ext.lower(),
        "Device": str(device),
        "Date": str(date_val),
        "Time": str(time_val),
        "GPS Coordinates": str(gps_val),
        "F-Number (Aperture)": str(f_number),
        "Focus / Focal Length": str(focal_length),
        "Brightness Value": str(brightness),
        "APEX Shutter Speed": str(apex_shutter),
        "APEX Aperture": str(apex_aperture),
        "ISO": str(iso),
        "Exposure Bias": str(exp_bias),
    }

    # Extract all other tags dynamically
    cleaned_extras = {}
    skip_keys = {"SourceFile", "Directory", "FilePermissions", "ExifToolVersion", "Thumbnail"}
    for k, v in raw_meta.items():
        if k in skip_keys:
            continue
        clean_key = k.split(":")[-1] if ":" in k else k
        if clean_key not in standard_data and str(v).lower() != "nan":
            cleaned_extras[clean_key] = str(v)

    return standard_data, cleaned_extras


def process_files_to_excel(uploaded_files):
    if not uploaded_files:
        return None, "❌ Please select at least one photo, video, or .zip archive."

    work_dir = tempfile.mkdtemp()
    excel_path = os.path.join(work_dir, "Metadata_Report.xlsx")

    # Step 1: Unpack any .zip archives or collect raw images
    target_media = []
    valid_exts = (".jpg", ".jpeg", ".png", ".webp", ".heic", ".mp4", ".mov")

    for f in uploaded_files:
        if f.lower().endswith(".zip"):
            extract_sub = tempfile.mkdtemp(dir=work_dir)
            with zipfile.ZipFile(f, "r") as zf:
                zf.extractall(extract_sub)
            for root, _, files in os.walk(extract_sub):
                for file_name in files:
                    if file_name.lower().endswith(valid_exts):
                        target_media.append(os.path.join(root, file_name))
        elif f.lower().endswith(valid_exts):
            target_media.append(f)

    if not target_media:
        return None, "❌ No valid photos or videos found in the selected files/archives."

    # Step 2: Read metadata across all media items
    records = []
    all_extra_keys = []

    for fpath in target_media:
        raw_meta = extract_all_exif(fpath)
        std_fields, extras = parse_photo_details(fpath, raw_meta)
        records.append({
            "source": fpath,
            "std": std_fields,
            "extras": extras
        })
        for key in extras.keys():
            if key not in all_extra_keys:
                all_extra_keys.append(key)

    standard_headers = list(records[0]["std"].keys())
    extra_headers = sorted(all_extra_keys)
    all_headers = standard_headers + extra_headers

    # Step 3: Create Styled Excel Document
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Metadata Report"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Calibri", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Write Header Row
    for col_num, header in enumerate(all_headers, 1):
        c = ws.cell(row=1, column=col_num, value=header)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center_align
        c.border = thin_border
    ws.row_dimensions[1].height = 28

    # Step 4: Write Data and Embed Thumbnails
    for row_idx, item in enumerate(records, start=2):
        src_path = item["source"]
        std = item["std"]
        extras = item["extras"]

        ws.row_dimensions[row_idx].height = 80

        # Fill Standard Fields
        for col_idx, col_name in enumerate(standard_headers, 1):
            val = std.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = cell_font
            cell.border = thin_border
            cell.alignment = center_align if col_name in ["Extension", "Date", "Time", "ISO", "GPS Coordinates"] else left_align

        # Fill Dynamic Extra Fields
        for col_offset, extra_col in enumerate(extra_headers, start=len(standard_headers) + 1):
            val = extras.get(extra_col, "")
            cell = ws.cell(row=row_idx, column=col_offset, value=val)
            cell.font = cell_font
            cell.border = thin_border
            cell.alignment = left_align

        # Resize & embed thumbnail in Column A
        try:
            with Image.open(src_path) as img:
                img_copy = img.copy()
                img_copy.thumbnail((95, 95))
                thumb_path = os.path.join(work_dir, f"thumb_{row_idx}.png")
                img_copy.save(thumb_path, "PNG")

                xl_img = OpenpyxlImage(thumb_path)
                ws.add_image(xl_img, f"A{row_idx}")
        except Exception:
            ws.cell(row=row_idx, column=1, value="No Preview").alignment = center_align

    # Set Column Widths
    ws.column_dimensions["A"].width = 16
    for col_idx in range(2, len(all_headers) + 1):
        col_letter = get_column_letter(col_idx)
        header_len = len(str(all_headers[col_idx - 1]))
        ws.column_dimensions[col_letter].width = max(header_len + 4, 15)

    ws.freeze_panes = "B2"
    wb.save(excel_path)

    summary_msg = (
        f"✅ Successfully processed {len(records)} photo(s)/video(s)!\n"
        f"📍 GPS, Timestamps, and Camera data extracted.\n"
        f"🖼️ Thumbnails embedded in Column A.\n"
        f"📊 Total metadata parameters: {len(all_headers)}\n\n"
        f"⬇️ Tap Download below to save your Excel report."
    )

    return excel_path, summary_msg


def count_files(files):
    if not files:
        return "No files selected."
    return f"### Selected: {len(files)} file(s) / archive(s)"


# =========================================================
# UI
# =========================================================
with gr.Blocks(title="Photo Metadata to Excel") as app:
    gr.Markdown(
        """
        # 📸 Photo & Video Metadata Extractor to Excel
        ### Full EXIF & GPS Extraction with Embedded Image Thumbnails

        * **Works with ZIP files or raw photos:** Compressing photos into a `.zip` in your phone's File Manager guarantees Android does not scrub GPS tags.
        * **Output:** Formatted `.xlsx` spreadsheet with embedded previews, device details, exact coordinates, and exposure metrics.
        """
    )

    files_input = gr.File(
        label="Select Photos / Videos or .ZIP Archive",
        file_count="multiple",
        type="filepath"
    )

    info_text = gr.Markdown("No files selected.")
    files_input.change(fn=count_files, inputs=files_input, outputs=info_text)

    run_btn = gr.Button("📑 GENERATE EXCEL WITH EMBEDDED IMAGES", variant="primary")
    excel_file = gr.File(label="Generated Excel Sheet (.xlsx)")
    status_box = gr.Textbox(label="Status & Summary", lines=6, interactive=False)

    run_btn.click(
        fn=process_files_to_excel,
        inputs=files_input,
        outputs=[excel_file, status_box]
    )

app.launch(share=True, debug=True)
