"""
verify_batch.py - Integrity and Quality Verification Script
"""
import os
import json
import pymupdf
from PIL import Image

output_dir = "output_certificates"
manifest_path = os.path.join(output_dir, "manifest.json")

print("==================================================================")
print("             OUTPUT ARTIFACT INTEGRITY VERIFICATION               ")
print("==================================================================")

if not os.path.exists(manifest_path):
    raise FileNotFoundError(f"Manifest not found at {manifest_path}")

with open(manifest_path, "r", encoding="utf-8") as f:
    manifest = json.load(f)

print(f"Manifest Timestamp    : {manifest.get('timestamp')}")
print(f"Total Records Ingested: {manifest.get('total_records')}")
print(f"Successful Generation : {manifest.get('successful')}")
print(f"Skipped Rows          : {manifest.get('skipped')}")
print(f"Failed Renderings     : {manifest.get('failed')}")
print(f"Processing Duration   : {manifest.get('elapsed_seconds')}s")
print("------------------------------------------------------------------")

for cert in manifest.get("certificates", []):
    name = cert["name"]
    course = cert["course"]
    cert_id = cert["cert_id"]
    pdf_filename = cert["files"]["pdf"]
    png_filename = cert["files"]["png"]

    pdf_path = os.path.join(output_dir, pdf_filename)
    png_path = os.path.join(output_dir, png_filename)

    # 1. PDF Verification
    assert os.path.exists(pdf_path), f"Missing PDF: {pdf_path}"
    pdf_size = os.path.getsize(pdf_path)
    assert pdf_size > 50000, f"PDF file size suspiciously small ({pdf_size} bytes)"

    doc = pymupdf.open(pdf_path)
    assert len(doc) == 1, f"Expected exactly 1 page in {pdf_filename}, found {len(doc)}"
    page = doc[0]
    rect = page.rect
    # Dimensions ~ 841.89 x 595.28 points
    assert abs(rect.width - 841.89) < 1.0, f"Unexpected width: {rect.width}"
    assert abs(rect.height - 595.28) < 1.0, f"Unexpected height: {rect.height}"

    text = page.get_text()
    assert cert_id in text, f"Certificate ID '{cert_id}' missing in PDF text stream"
    doc.close()

    # 2. PNG Verification
    assert os.path.exists(png_path), f"Missing PNG: {png_path}"
    png_size = os.path.getsize(png_path)
    assert png_size > 100000, f"PNG file size suspiciously small ({png_size} bytes)"

    with Image.open(png_path) as img:
        w, h = img.size
        # 300 DPI A4 landscape ~ 3508 x 2480-2481 pixels
        assert abs(w - 3508) <= 2 and abs(h - 2480) <= 2, f"Unexpected resolution: {w}x{h}"
        assert img.format == "PNG", f"Expected PNG format, got {img.format}"

    print(f"  [VERIFIED] {name:<38} | PDF: {pdf_size//1024:>4} KB | PNG: {w}x{h} ({png_size//1024:>4} KB)")

print("------------------------------------------------------------------")
print("VERIFICATION RESULT: ALL ARTIFACTS PASSED 100% VALIDATION!")
print("==================================================================")
