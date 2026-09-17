"""
mail_merge.py - High-Resolution Certificate Generation & Mail-Merge Pipeline

Author: Automation Engineering
Description:
    Processes recipient datasets (CSV/Excel), applies certificate template styling,
    dynamically prevents text overflow via font auto-scaling, and renders vector PDFs
    and high-resolution 300-DPI PNG certificates with comprehensive error logging.
"""

import os
import sys
import re
import json
import time
import math
import hashlib
import logging
import argparse
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

import pandas as pd
import pymupdf
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


def setup_logger(log_file: Optional[str] = None, verbose: bool = False) -> logging.Logger:
    """Configures dual console and file logging."""
    logger = logging.getLogger("CertificateMailMerge")
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.handlers.clear()

    c_handler = logging.StreamHandler(sys.stdout)
    c_handler.setLevel(logging.DEBUG if verbose else logging.INFO)
    c_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(c_handler)

    if log_file:
        os.makedirs(os.path.dirname(os.path.abspath(log_file)), exist_ok=True)
        f_handler = logging.FileHandler(log_file, encoding="utf-8")
        f_handler.setLevel(logging.DEBUG)
        f_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s"))
        logger.addHandler(f_handler)

    return logger


class FontManager:
    """Discovers and registers TrueType fonts with fallback to standard Type 1 fonts."""

    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.fonts: Dict[str, str] = {}
        self._register_fonts()

    def _register_fonts(self):
        font_candidates = {
            "Georgia": [r"C:\Windows\Fonts\georgia.ttf", "/usr/share/fonts/truetype/msttcorefonts/georgia.ttf"],
            "Georgia-Bold": [r"C:\Windows\Fonts\georgiab.ttf", "/usr/share/fonts/truetype/msttcorefonts/georgiab.ttf"],
            "Georgia-Italic": [r"C:\Windows\Fonts\georgiai.ttf", "/usr/share/fonts/truetype/msttcorefonts/georgiai.ttf"],
            "Arial": [r"C:\Windows\Fonts\arial.ttf", "/Library/Fonts/Arial.ttf"],
            "Arial-Bold": [r"C:\Windows\Fonts\arialbd.ttf", "/Library/Fonts/Arial Bold.ttf"],
            "Arial-Italic": [r"C:\Windows\Fonts\ariali.ttf", "/Library/Fonts/Arial Italic.ttf"],
        }
        registered = 0
        for name, paths in font_candidates.items():
            for p in paths:
                if os.path.exists(p):
                    try:
                        pdfmetrics.registerFont(TTFont(name, p))
                        self.fonts[name] = name
                        registered += 1
                        break
                    except Exception as e:
                        self.logger.debug(f"Failed to register font {name}: {e}")
        self.logger.debug(f"Registered {registered} TrueType fonts.")

    def get_font(self, font_type: str) -> str:
        mapping = {
            "serif_regular": "Georgia" if "Georgia" in self.fonts else "Times-Roman",
            "serif_bold": "Georgia-Bold" if "Georgia-Bold" in self.fonts else "Times-Bold",
            "serif_italic": "Georgia-Italic" if "Georgia-Italic" in self.fonts else "Times-Italic",
            "sans_regular": "Arial" if "Arial" in self.fonts else "Helvetica",
            "sans_bold": "Arial-Bold" if "Arial-Bold" in self.fonts else "Helvetica-Bold",
            "sans_italic": "Arial-Italic" if "Arial-Italic" in self.fonts else "Helvetica-Oblique",
        }
        return mapping.get(font_type, "Helvetica")


def sanitize_filename(name: str, cert_id: Optional[str] = None) -> str:
    """Generates a safe filename without special characters or illegal path chars."""
    normalized = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    clean_name = re.sub(r"[^a-zA-Z0-9_\-\.]+", "_", normalized).strip("_")
    if not clean_name:
        clean_name = "Recipient"

    if cert_id:
        clean_id = re.sub(r"[^a-zA-Z0-9_\-]+", "", cert_id).strip()
        filename = f"Certificate_{clean_name}_{clean_id}"
    else:
        filename = f"Certificate_{clean_name}"

    return filename[:120]

class CertificateRenderer:
    """Renders high-resolution vector PDF certificates using ReportLab."""

    def __init__(self, config_path: str, font_manager: FontManager, logger: logging.Logger):
        self.logger = logger
        self.fonts = font_manager
        self.config = self._load_config(config_path)

        page_cfg = self.config.get("page", {})
        self.width = float(page_cfg.get("width_pt", 841.89))
        self.height = float(page_cfg.get("height_pt", 595.28))

        theme = self.config.get("theme", {})
        self.c_bg = HexColor(theme.get("background_color", "#FCFDFD"))
        self.c_primary = HexColor(theme.get("primary_color", "#0B192C"))
        self.c_gold = HexColor(theme.get("accent_gold", "#C59B27"))
        self.c_light_gold = HexColor(theme.get("light_gold", "#E8D595"))
        self.c_dark = HexColor(theme.get("text_dark", "#1E293B"))
        self.c_muted = HexColor(theme.get("text_muted", "#64748B"))

    def _load_config(self, config_path: str) -> Dict[str, Any]:
        if not os.path.exists(config_path):
            self.logger.warning(f"Config '{config_path}' not found! Using standard defaults.")
            return {}
        with open(config_path, "r", encoding="utf-8-sig") as f:
            return json.load(f)

    def _draw_borders_and_decorations(self, c: canvas.Canvas):
        """Draws ornate double borders, corner flourishes, and background."""
        c.setFillColor(self.c_bg)
        c.rect(0, 0, self.width, self.height, stroke=0, fill=1)

        outer_margin = 24
        c.setStrokeColor(self.c_primary)
        c.setLineWidth(4)
        c.rect(outer_margin, outer_margin, self.width - 2 * outer_margin, self.height - 2 * outer_margin, stroke=1, fill=0)

        inner_margin = 32
        c.setStrokeColor(self.c_gold)
        c.setLineWidth(1.5)
        c.rect(inner_margin, inner_margin, self.width - 2 * inner_margin, self.height - 2 * inner_margin, stroke=1, fill=0)

        # Corner diamond accents
        corner_coords = [
            (inner_margin + 6, inner_margin + 6),
            (self.width - inner_margin - 6, inner_margin + 6),
            (inner_margin + 6, self.height - inner_margin - 6),
            (self.width - inner_margin - 6, self.height - inner_margin - 6),
        ]
        c.setFillColor(self.c_gold)
        diamond_radius = 4
        for cx, cy in corner_coords:
            p = c.beginPath()
            p.moveTo(cx, cy + diamond_radius)
            p.lineTo(cx + diamond_radius, cy)
            p.lineTo(cx, cy - diamond_radius)
            p.lineTo(cx - diamond_radius, cy)
            p.close()
            c.drawPath(p, stroke=0, fill=1)

        # Corner brackets
        c.setLineWidth(1.2)
        blen = 22
        off = 36
        c.line(off, self.height - off, off + blen, self.height - off)
        c.line(off, self.height - off, off, self.height - off - blen)
        c.line(self.width - off, self.height - off, self.width - off - blen, self.height - off)
        c.line(self.width - off, self.height - off, self.width - off, self.height - off - blen)
        c.line(off, off, off + blen, off)
        c.line(off, off, off, off + blen)
        c.line(self.width - off, off, self.width - off - blen, off)
        c.line(self.width - off, off, self.width - off, off + blen)

    def _draw_official_seal(self, c: canvas.Canvas, cx: float, cy: float, radius: float = 38):
        """Draws official gold rosette vector with ribbon and 5-point star."""
        c.setFillColor(self.c_primary)
        for offset_x in [-16, 6]:
            p = c.beginPath()
            p.moveTo(cx + offset_x, cy - radius + 8)
            p.lineTo(cx + offset_x - 8, cy - radius - 24)
            p.lineTo(cx + offset_x + 2, cy - radius - 18)
            p.lineTo(cx + offset_x + 12, cy - radius - 24)
            p.lineTo(cx + offset_x + 10, cy - radius + 8)
            p.close()
            c.drawPath(p, stroke=0, fill=1)

        # Scalloped outer rosette
        num_points = 32
        rosette = c.beginPath()
        for i in range(num_points):
            angle = (i * 2 * math.pi) / num_points
            r = radius if (i % 2 == 0) else radius - 3.5
            x = cx + r * math.cos(angle)
            y = cy + r * math.sin(angle)
            if i == 0:
                rosette.moveTo(x, y)
            else:
                rosette.lineTo(x, y)
        rosette.close()
        c.setFillColor(self.c_gold)
        c.drawPath(rosette, stroke=0, fill=1)

        c.setFillColor(self.c_bg)
        c.setStrokeColor(self.c_gold)
        c.setLineWidth(1.5)
        c.circle(cx, cy, radius - 6, stroke=1, fill=1)

        # Center star
        star_radius = 8.5
        star = c.beginPath()
        for i in range(5):
            angle = (i * 4 * math.pi) / 5 - math.pi / 2
            ox = cx + star_radius * math.cos(angle)
            oy = cy + star_radius * math.sin(angle)
            if i == 0:
                star.moveTo(ox, oy)
            else:
                star.lineTo(ox, oy)
        star.close()
        c.setFillColor(self.c_gold)
        c.drawPath(star, stroke=0, fill=1)

        c.setFont(self.fonts.get_font("sans_bold"), 5.5)
        c.setFillColor(self.c_primary)
        c.drawCentredString(cx, cy + 14, "OFFICIAL SEAL")
        c.drawCentredString(cx, cy - 18, "EXCELLENCE")

    def _fit_text(self, c: canvas.Canvas, text: str, font_name: str, max_size: float, min_size: float, max_width: float) -> Tuple[float, List[str]]:
        """Calculates optimal font size or wraps text to prevent overflow."""
        size = max_size
        while size >= min_size:
            if c.stringWidth(text, font_name, size) <= max_width:
                return size, [text]
            size -= 0.5

        words = text.split()
        if len(words) <= 1:
            return min_size, [text]

        half = len(words) // 2
        line1 = " ".join(words[:half])
        line2 = " ".join(words[half:])

        wrapped_size = min_size
        while wrapped_size >= 10:
            w1 = c.stringWidth(line1, font_name, wrapped_size)
            w2 = c.stringWidth(line2, font_name, wrapped_size)
            if max(w1, w2) <= max_width:
                return wrapped_size, [line1, line2]
            wrapped_size -= 0.5

        return 10.0, [line1, line2]

    def render_certificate(self, record: Dict[str, Any], output_pdf_path: str):
        """Renders single personalized vector certificate PDF."""
        c = canvas.Canvas(output_pdf_path, pagesize=(self.width, self.height))

        # 1. Background & borders
        self._draw_borders_and_decorations(c)

        center_x = self.width / 2.0
        max_content_w = self.width * 0.76

        # 2. Institution Header
        c.setFont(self.fonts.get_font("sans_bold"), 10)
        c.setFillColor(self.c_muted)
        org = self.config.get("branding", {}).get("organization", "INTERNATIONAL INSTITUTE OF TECHNOLOGY & DATA SCIENCES")
        c.drawCentredString(center_x, self.height - 70, org)

        c.setStrokeColor(self.c_gold)
        c.setLineWidth(1)
        c.line(center_x - 130, self.height - 78, center_x + 130, self.height - 78)

        # 3. Main Title
        c.setFont(self.fonts.get_font("serif_bold"), 30)
        c.setFillColor(self.c_primary)
        title = self.config.get("branding", {}).get("certificate_title", "CERTIFICATE OF ACHIEVEMENT")
        c.drawCentredString(center_x, self.height - 118, title)

        # 4. Presentation Line
        c.setFont(self.fonts.get_font("sans_bold"), 10)
        c.setFillColor(self.c_gold)
        pres_line = self.config.get("branding", {}).get("presentation_line", "THIS IS PROUDLY PRESENTED TO")
        c.drawCentredString(center_x, self.height - 148, pres_line)

        # 5. Dynamic Recipient Name (auto-scaled)
        name = str(record.get("name", "")).strip()
        f_serif_bold = self.fonts.get_font("serif_bold")
        scale_cfg = self.config.get("text_scaling", {})
        n_size, n_lines = self._fit_text(c, name, f_serif_bold,
                                         float(scale_cfg.get("name_max_pt", 34)),
                                         float(scale_cfg.get("name_min_pt", 16)),
                                         max_content_w)

        c.setFont(f_serif_bold, n_size)
        c.setFillColor(self.c_dark)

        if len(n_lines) == 1:
            name_y = self.height - 200
            c.drawCentredString(center_x, name_y, n_lines[0])
            w = c.stringWidth(n_lines[0], f_serif_bold, n_size)
            bar_w = min(max(w + 50, 220), max_content_w)
            c.setStrokeColor(self.c_gold)
            c.setLineWidth(1.5)
            c.line(center_x - bar_w / 2, name_y - 8, center_x + bar_w / 2, name_y - 8)

            c.setFillColor(self.c_gold)
            p = c.beginPath()
            p.moveTo(center_x, name_y - 5)
            p.lineTo(center_x + 3.5, name_y - 8)
            p.lineTo(center_x, name_y - 11)
            p.lineTo(center_x - 3.5, name_y - 8)
            p.close()
            c.drawPath(p, stroke=0, fill=1)
        else:
            name_y = self.height - 192
            c.drawCentredString(center_x, name_y, n_lines[0])
            c.drawCentredString(center_x, name_y - n_size - 4, n_lines[1])
            total_w = max(c.stringWidth(n_lines[0], f_serif_bold, n_size), c.stringWidth(n_lines[1], f_serif_bold, n_size))
            bar_w = min(total_w + 50, max_content_w)
            c.setStrokeColor(self.c_gold)
            c.setLineWidth(1.5)
            c.line(center_x - bar_w / 2, name_y - n_size - 14, center_x + bar_w / 2, name_y - n_size - 14)

        # 6. Completion Reason
        c.setFont(self.fonts.get_font("serif_italic"), 12)
        c.setFillColor(self.c_muted)
        comp_text = self.config.get("branding", {}).get("completion_line", "for successfully fulfilling all rigorous academic requirements and demonstrating mastery in")
        c.drawCentredString(center_x, self.height - 250, comp_text)

        # 7. Course Title (auto-scaled)
        course = str(record.get("course", "")).strip()
        f_sans_bold = self.fonts.get_font("sans_bold")
        c_size, c_lines = self._fit_text(c, course, f_sans_bold,
                                         float(scale_cfg.get("course_max_pt", 22)),
                                         float(scale_cfg.get("course_min_pt", 13)),
                                         max_content_w)
        c.setFont(f_sans_bold, c_size)
        c.setFillColor(self.c_primary)

        if len(c_lines) == 1:
            c.drawCentredString(center_x, self.height - 285, c_lines[0])
        else:
            c.drawCentredString(center_x, self.height - 280, c_lines[0])
            c.drawCentredString(center_x, self.height - 280 - c_size - 4, c_lines[1])

        # 8. Signatures & Official Seal
        sig_cfg = self.config.get("signatures", {})
        sig_l = sig_cfg.get("left", {"signatory_name": "Dr. Aris Thorne", "signatory_title": "Chief Academic Officer"})
        sig_r = sig_cfg.get("right", {"signatory_name": "Prof. H. R. Vance", "signatory_title": "Director of Education"})

        sig_y = 135
        sig_w = 160

        # Left Signatory
        lx = self.width * 0.24
        c.setFont(self.fonts.get_font("serif_italic"), 15)
        c.setFillColor(self.c_primary)
        c.drawCentredString(lx, sig_y + 16, sig_l.get("signatory_name", ""))
        c.setStrokeColor(self.c_primary)
        c.setLineWidth(1)
        c.line(lx - sig_w / 2, sig_y + 8, lx + sig_w / 2, sig_y + 8)
        c.setFont(self.fonts.get_font("sans_bold"), 9.5)
        c.setFillColor(self.c_dark)
        c.drawCentredString(lx, sig_y - 6, sig_l.get("signatory_name", ""))
        c.setFont(self.fonts.get_font("sans_regular"), 8)
        c.setFillColor(self.c_muted)
        c.drawCentredString(lx, sig_y - 18, sig_l.get("signatory_title", ""))

        # Center Official Seal
        self._draw_official_seal(c, center_x, sig_y + 18, radius=38)

        # Right Signatory
        rx = self.width * 0.76
        r_name = record.get("instructor") or sig_r.get("signatory_name", "Prof. H. R. Vance")
        r_title = "Lead Course Instructor" if record.get("instructor") else sig_r.get("signatory_title", "Director of Education")
        c.setFont(self.fonts.get_font("serif_italic"), 15)
        c.setFillColor(self.c_primary)
        c.drawCentredString(rx, sig_y + 16, r_name)
        c.setStrokeColor(self.c_primary)
        c.setLineWidth(1)
        c.line(rx - sig_w / 2, sig_y + 8, rx + sig_w / 2, sig_y + 8)
        c.setFont(self.fonts.get_font("sans_bold"), 9.5)
        c.setFillColor(self.c_dark)
        c.drawCentredString(rx, sig_y - 6, r_name)
        c.setFont(self.fonts.get_font("sans_regular"), 8)
        c.setFillColor(self.c_muted)
        c.drawCentredString(rx, sig_y - 18, r_title)

        # 9. Meta Footer (Safely within borders)
        c.setFont(self.fonts.get_font("sans_regular"), 8)
        c.setFillColor(self.c_muted)
        date_str = str(record.get("date", "")).strip()
        cert_id = str(record.get("cert_id", "")).strip()

        footer_y = 46
        footer_inset = 60
        c.drawString(footer_inset, footer_y, f"Issued: {date_str}")
        prefix = self.config.get("branding", {}).get("verification_prefix", "Official Credential ID: ")
        c.drawCentredString(center_x, footer_y, f"{prefix}{cert_id}")
        h = hashlib.sha256(f"{name}:{course}:{cert_id}".encode()).hexdigest()[:12].upper()
        c.drawRightString(self.width - footer_inset, footer_y, f"Security Hash: {h}")

        c.save()

class DataProcessor:
    """Parses and standardizes recipient datasets from CSV or Excel."""

    COLUMN_SYNONYMS = {
        "name": ["name", "full name", "fullname", "recipient", "recipient name", "student", "student name"],
        "course": ["course", "course name", "coursename", "program", "program name", "title", "track"],
        "date": ["date", "completion date", "issue date", "issued date", "date of completion", "graduation date"],
        "cert_id": ["cert_id", "certificate id", "cert id", "id", "credential id", "certificate_id"],
        "instructor": ["instructor", "teacher", "professor", "signatory", "director", "instructor name"]
    }

    def __init__(self, file_path: str, logger: logging.Logger):
        self.file_path = file_path
        self.logger = logger

    def load_records(self) -> List[Dict[str, Any]]:
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"Data source file not found: {self.file_path}")

        ext = Path(self.file_path).suffix.lower()
        if ext == ".csv":
            df = pd.read_csv(self.file_path, dtype=str)
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(self.file_path, dtype=str)
        else:
            raise ValueError(f"Unsupported file format '{ext}'. Must be .csv or .xlsx/.xls")

        self.logger.info(f"Loaded {len(df)} records from '{self.file_path}'.")

        normalized = []
        for idx, row in df.iterrows():
            rec = {"__row_index__": idx + 1}
            for col in df.columns:
                val = "" if pd.isna(row[col]) else str(row[col]).strip()
                norm_col = col.strip().lower()

                matched = False
                for standard_key, synonyms in self.COLUMN_SYNONYMS.items():
                    if norm_col in synonyms:
                        rec[standard_key] = val
                        matched = True
                        break
                if not matched:
                    rec[norm_col] = val

            normalized.append(rec)
        return normalized


class CertificatePipeline:
    """Coordinates batch mail-merge, formatting, and file export."""

    def __init__(self, data_file: str, template_config: str, output_dir: str,
                 export_format: str = "both", dpi: int = 300, log_file: Optional[str] = None, verbose: bool = False):
        self.data_file = data_file
        self.template_config = template_config
        self.output_dir = output_dir
        self.export_format = export_format.lower()
        self.dpi = dpi
        self.logger = setup_logger(log_file, verbose)

        self.fonts = FontManager(self.logger)
        self.renderer = CertificateRenderer(self.template_config, self.fonts, self.logger)
        self.processor = DataProcessor(self.data_file, self.logger)

        os.makedirs(self.output_dir, exist_ok=True)

    def validate_record(self, record: Dict[str, Any], row_idx: int) -> Tuple[bool, Optional[str]]:
        name = record.get("name", "").strip()
        course = record.get("course", "").strip()

        if not name:
            return False, f"Row {row_idx}: Missing required field 'Full Name' / 'Name'."
        if not course:
            return False, f"Row {row_idx}: Missing required field 'Course Name' / 'Course'."

        if not record.get("date"):
            record["date"] = datetime.now().strftime("%Y-%m-%d")
            self.logger.debug(f"Row {row_idx}: Auto-assigned date {record['date']}.")

        if not record.get("cert_id"):
            token = hashlib.md5(f"{name}:{course}:{time.time()}".encode()).hexdigest()[:8].upper()
            record["cert_id"] = f"CERT-GEN-{token}"
            self.logger.debug(f"Row {row_idx}: Auto-assigned ID {record['cert_id']}.")

        return True, None

    def export_png(self, pdf_path: str, png_path: str):
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        zoom = self.dpi / 72.0
        mat = pymupdf.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        pix.save(png_path)
        doc.close()

    def run(self, dry_run: bool = False) -> Dict[str, Any]:
        start = time.time()
        self.logger.info(f"Starting Certificate Batch Pipeline (Dry Run: {dry_run})")
        self.logger.info(f"Export Format: {self.export_format.upper()} (PNG DPI: {self.dpi})")
        self.logger.info(f"Output Directory: {os.path.abspath(self.output_dir)}")

        records = self.processor.load_records()
        manifest = []
        successful, skipped, failed = 0, 0, 0
        used_names = set()

        for rec in records:
            row_idx = rec.get("__row_index__", "?")
            valid, err = self.validate_record(rec, row_idx)

            if not valid:
                self.logger.warning(f"[SKIPPED] {err}")
                skipped += 1
                continue

            if dry_run:
                self.logger.info(f"[DRY RUN VALID] Row {row_idx}: '{rec['name']}' -> '{rec['course']}'")
                successful += 1
                continue

            base_name = sanitize_filename(rec["name"], rec.get("cert_id"))
            target_name = base_name
            dup_cnt = 1
            while target_name in used_names:
                target_name = f"{base_name}_{dup_cnt}"
                dup_cnt += 1
            used_names.add(target_name)

            pdf_out = os.path.join(self.output_dir, f"{target_name}.pdf")
            png_out = os.path.join(self.output_dir, f"{target_name}.png")

            try:
                self.renderer.render_certificate(rec, pdf_out)
                pdf_size = os.path.getsize(pdf_out)

                png_size = None
                if self.export_format in ("png", "both"):
                    self.export_png(pdf_out, png_out)
                    png_size = os.path.getsize(png_out)

                if self.export_format == "png" and os.path.exists(pdf_out):
                    os.remove(pdf_out)

                successful += 1
                msg_parts = []
                if self.export_format in ("pdf", "both"):
                    msg_parts.append(f"PDF: {pdf_size // 1024} KB")
                if self.export_format in ("png", "both"):
                    msg_parts.append(f"PNG: {png_size // 1024} KB")

                self.logger.info(f"[SUCCESS] Row {row_idx}: {rec['name']} -> {target_name} ({', '.join(msg_parts)})")

                manifest.append({
                    "row": row_idx,
                    "name": rec["name"],
                    "course": rec["course"],
                    "cert_id": rec["cert_id"],
                    "date": rec["date"],
                    "files": {
                        "pdf": os.path.basename(pdf_out) if self.export_format in ("pdf", "both") else None,
                        "png": os.path.basename(png_out) if self.export_format in ("png", "both") else None
                    }
                })

            except Exception as e:
                failed += 1
                self.logger.error(f"[ERROR] Row {row_idx}: Rendering failed for '{rec.get('name')}': {e}", exc_info=True)

        elapsed = time.time() - start

        if not dry_run:
            manifest_path = os.path.join(self.output_dir, "manifest.json")
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump({
                    "timestamp": datetime.now().isoformat(),
                    "total_records": len(records),
                    "successful": successful,
                    "skipped": skipped,
                    "failed": failed,
                    "elapsed_seconds": round(elapsed, 2),
                    "certificates": manifest
                }, f, indent=2)

        self.logger.info("=" * 65)
        self.logger.info("BATCH GENERATION SUMMARY")
        self.logger.info(f"  Total Ingested Records : {len(records)}")
        self.logger.info(f"  Successfully Generated  : {successful}")
        self.logger.info(f"  Skipped (Invalid/Blank): {skipped}")
        self.logger.info(f"  Failed (Exceptions)    : {failed}")
        self.logger.info(f"  Elapsed Processing Time : {elapsed:.2f}s")
        self.logger.info(f"  Manifest Output Path   : {manifest_path}")
        self.logger.info("=" * 65)

        return {"total": len(records), "successful": successful, "skipped": skipped, "failed": failed}


def main():
    parser = argparse.ArgumentParser(description="High-Resolution Certificate Mail-Merge Pipeline")
    parser.add_argument("-d", "--data", default="recipients.csv", help="Recipient CSV or Excel file")
    parser.add_argument("-t", "--template", default="templates/template_config.json", help="Template JSON configuration")
    parser.add_argument("-o", "--output", default="output_certificates", help="Output directory")
    parser.add_argument("-f", "--format", choices=["pdf", "png", "both"], default="both", help="Export file format")
    parser.add_argument("--dpi", type=int, default=300, help="DPI resolution for PNG exports")
    parser.add_argument("--log", default="logs/generation.log", help="Path to log file")
    parser.add_argument("--dry-run", action="store_true", help="Validate records without generating files")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose debug logging")
    parser.add_argument("-a", "--agent", "--interactive", action="store_true", help="Launch interactive Certificate Generation Agent")

    args = parser.parse_args()

    if args.agent:
        import certificate_agent
        certificate_agent.main([])
        return

    pipeline = CertificatePipeline(
        data_file=args.data,
        template_config=args.template,
        output_dir=args.output,
        export_format=args.format,
        dpi=args.dpi,
        log_file=args.log,
        verbose=args.verbose
    )

    results = pipeline.run(dry_run=args.dry_run)
    if results["failed"] > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()

