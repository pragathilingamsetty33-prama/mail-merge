"""
certificate_agent.py - User-Friendly Interactive Certificate Generation Agent

A non-technical, interactive assistant for automated certificate generation.
- Asks user for CSV/Excel file (with auto-detection, browse dialog, and drag-and-drop support)
- Reads and normalizes recipient records
- Validates required fields (Name and Course)
- Explains which rows are invalid and why
- Accidental overwrite protection (Skip, Version, or Confirmed Overwrite)
- Renders vector PDF and 300-DPI PNG certificates
- Saves outputs into output_certificates/
- Displays a clean final execution summary
"""

import os
import sys

# Ensure robust console output encoding on Windows
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import glob
import time
import json
import hashlib
import argparse
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Tuple, Optional

# Re-use proven core rendering and data parsing from mail_merge.py
from mail_merge import (
    CertificateRenderer,
    DataProcessor,
    FontManager,
    sanitize_filename,
    setup_logger
)
import pymupdf


class CertificateAgent:
    """Friendly interactive agent guiding users through certificate generation."""

    def __init__(
        self,
        template_config_path: str = "templates/template_config.json",
        output_dir: str = "output_certificates",
        dpi: int = 300,
        log_file: str = "logs/generation.log"
    ):
        self.template_config_path = template_config_path
        self.output_dir = output_dir
        self.dpi = dpi
        os.makedirs(self.output_dir, exist_ok=True)

        self.logger = setup_logger(log_file=log_file, verbose=False)
        self.font_manager = FontManager(self.logger)
        self.renderer = CertificateRenderer(self.template_config_path, self.font_manager, self.logger)

    def print_banner(self):
        """Prints a friendly welcome banner."""
        print("=" * 68)
        print("          AUTOMATED CERTIFICATE GENERATION AGENT")
        print("=" * 68)
        print(" Welcome! This assistant creates high-resolution vector PDF and 300-DPI")
        print(" PNG certificates from your CSV or Excel recipient list.")
        print("=" * 68)

    def detect_workspace_data_files(self) -> List[str]:
        """Finds CSV and Excel files in current working directory."""
        candidates = []
        for ext in ("*.csv", "*.xlsx", "*.xls"):
            candidates.extend(glob.glob(ext))
        def priority(f):
            f_lower = f.lower()
            if "recipient" in f_lower:
                return (0, f_lower)
            return (1, f_lower)
        candidates.sort(key=priority)
        return candidates

    def prompt_recipient_file(self) -> Optional[str]:
        """
        Asks user for the recipient CSV/Excel file with smart suggestions,
        drag-and-drop parsing, and file dialog support.
        """
        files = self.detect_workspace_data_files()

        print("\n[Step 1 of 4] Select Recipient Data File")
        print("-" * 68)

        if files:
            print("Found the following spreadsheet(s) in this folder:")
            for i, f in enumerate(files, 1):
                try:
                    size_kb = os.path.getsize(f) // 1024
                    print(f"  [{i}] {f} ({size_kb} KB)")
                except Exception:
                    print(f"  [{i}] {f}")
            default_file = files[0]
            print(f"\nPress Enter to use default: [{default_file}]")
        else:
            default_file = "recipients.csv"
            print(f"Enter path to your CSV or Excel file (default: {default_file})")

        print("Options: Enter a number, paste a path, type 'B' to browse, or 'Q' to quit.")

        while True:
            try:
                user_input = input("\n>> File path or selection: ").strip()
            except (KeyboardInterrupt, EOFError):
                print("\nOperation cancelled by user.")
                return None

            if not user_input:
                if os.path.exists(default_file):
                    return default_file
                print(f"[!] Default file '{default_file}' was not found. Please provide a path.")
                continue

            # Quit option
            if user_input.lower() in ("q", "quit", "exit"):
                print("Exiting certificate agent. Goodbye!")
                return None

            # File dialog browse option
            if user_input.lower() in ("b", "browse"):
                print("Opening file browser window...")
                try:
                    import tkinter as tk
                    from tkinter import filedialog
                    root = tk.Tk()
                    root.withdraw()
                    root.attributes("-topmost", True)
                    chosen = filedialog.askopenfilename(
                        title="Select Recipient Spreadsheet",
                        filetypes=[
                            ("Spreadsheets (CSV, Excel)", "*.csv;*.xlsx;*.xls"),
                            ("CSV files (*.csv)", "*.csv"),
                            ("Excel files (*.xlsx;*.xls)", "*.xlsx;*.xls"),
                            ("All files (*.*)", "*.*")
                        ]
                    )
                    root.destroy()
                    if chosen:
                        print(f"Selected: {chosen}")
                        return chosen
                    else:
                        print("No file selected in browser. Please try again.")
                        continue
                except Exception as e:
                    print(f"[!] Could not open GUI file browser ({e}). Please type the path.")
                    continue

            # Number selection from listed files
            if user_input.isdigit() and files:
                num = int(user_input)
                if 1 <= num <= len(files):
                    return files[num - 1]
                else:
                    print(f"[!] Please enter a number between 1 and {len(files)}.")
                    continue

            # Clean path from drag-and-drop (handles quotes or leading '&' on Windows)
            clean_path = user_input.strip(' "\'').lstrip('& ')
            if clean_path.startswith('"') and clean_path.endswith('"'):
                clean_path = clean_path[1:-1]

            if not os.path.exists(clean_path):
                print(f"[!] File not found: '{clean_path}'. Please check the path and try again.")
                continue

            ext = Path(clean_path).suffix.lower()
            if ext not in (".csv", ".xlsx", ".xls"):
                print(f"[!] Unsupported format '{ext}'. Please select a .csv, .xlsx, or .xls file.")
                continue

            return clean_path

    def inspect_and_validate(
        self,
        file_path: str
    ) -> Tuple[List[Dict[str, Any]], List[Tuple[int, Dict[str, Any], str]]]:
        """
        Loads all records using DataProcessor and validates required fields.
        Returns:
            (valid_records, invalid_records)
        """
        processor = DataProcessor(file_path, self.logger)
        records = processor.load_records()

        valid_records = []
        invalid_records = []

        for rec in records:
            row_idx = rec.get("__row_index__", "?")
            name = str(rec.get("name", "")).strip()
            course = str(rec.get("course", "")).strip()

            missing = []
            if not name:
                missing.append("'Full Name' / 'Name'")
            if not course:
                missing.append("'Course Name' / 'Course'")

            if missing:
                reason = f"Missing required field(s): {', '.join(missing)}"
                invalid_records.append((row_idx, rec, reason))
            else:
                # Auto-assign date if missing
                if not rec.get("date"):
                    rec["date"] = datetime.now().strftime("%Y-%m-%d")
                # Auto-assign certificate ID if missing
                if not rec.get("cert_id"):
                    token = hashlib.md5(f"{name}:{course}:{time.time()}".encode()).hexdigest()[:8].upper()
                    rec["cert_id"] = f"CERT-GEN-{token}"
                valid_records.append(rec)

        return valid_records, invalid_records

    def display_validation_report(
        self,
        file_path: str,
        valid_records: List[Dict[str, Any]],
        invalid_records: List[Tuple[int, Dict[str, Any], str]]
    ) -> bool:
        """Displays data validation status and details of invalid rows."""
        total = len(valid_records) + len(invalid_records)

        print("\n[Step 2 of 4] Recipient Data Validation Report")
        print("-" * 68)
        print(f"  Source File           : {os.path.basename(file_path)}")
        print(f"  Total Ingested Rows   : {total}")
        print(f"  [OK] Valid Recipients : {len(valid_records)}")
        print(f"  [!] Incomplete/Skipped: {len(invalid_records)}")
        print("-" * 68)

        if invalid_records:
            print("\n[!] THE FOLLOWING ROWS ARE INCOMPLETE AND WILL BE SKIPPED:")
            for row_idx, rec, reason in invalid_records:
                print(f"\n  * Row {row_idx}: {reason}")
                details = []
                if rec.get("name"):
                    details.append(f"Name='{rec['name']}'")
                if rec.get("course"):
                    details.append(f"Course='{rec['course']}'")
                if rec.get("date"):
                    details.append(f"Date='{rec['date']}'")
                if rec.get("cert_id"):
                    details.append(f"ID='{rec['cert_id']}'")
                if details:
                    print(f"    Existing data in row: {', '.join(details)}")
                else:
                    print("    Existing data in row: [Blank row or missing required columns]")
            print("\n" + "-" * 68)

        if not valid_records:
            print("[!] No valid records found in this file. Cannot proceed with generation.")
            return False

        return True

    def check_and_resolve_collisions(
        self,
        valid_records: List[Dict[str, Any]],
        default_choice: Optional[str] = None
    ) -> Tuple[str, List[Tuple[Dict[str, Any], str, str, str]]]:
        """
        Accidental Overwrite Protection (Requirement 9):
        Scans output_certificates/ for any existing files matching valid recipients.
        Prompts user if collisions exist, defaulting to Safe Skip mode.
        Returns:
            (mode, planned_tasks)
            mode in ('skip', 'version', 'overwrite')
        """
        planned_tasks = []
        collisions = []

        print("\n[Step 3 of 4] Checking Existing Certificates & Overwrite Safety")
        print("-" * 68)

        for rec in valid_records:
            base_name = sanitize_filename(rec["name"], rec.get("cert_id"))
            pdf_path = os.path.join(self.output_dir, f"{base_name}.pdf")
            png_path = os.path.join(self.output_dir, f"{base_name}.png")

            exists = os.path.exists(pdf_path) or os.path.exists(png_path)
            if exists:
                collisions.append((rec, base_name, pdf_path, png_path))
            planned_tasks.append((rec, base_name, pdf_path, png_path))

        if not collisions:
            print("[OK] No existing certificates found in output folder. Ready for clean generation.")
            return "overwrite", planned_tasks

        # Collisions detected: Protect existing files!
        print(f"[*] Found {len(collisions)} certificate(s) already in '{self.output_dir}':")
        for rec, base_name, _, _ in collisions[:5]:
            print(f"   - {base_name}.pdf (.png)")
        if len(collisions) > 5:
            print(f"   - ... and {len(collisions) - 5} more.")

        if default_choice in ("skip", "version", "overwrite"):
            return default_choice, planned_tasks

        print("\nTo prevent accidental file loss, choose how to handle existing certificates:")
        print("  [1] Skip existing       - Keep previous certificates, do not recreate [DEFAULT - SAFE]")
        print("  [2] Create new version  - Keep previous files, save new ones as _v2, _v3 [SAFE]")
        print("  [3] Overwrite existing  - Re-generate and replace previous files [REQUIRES CONFIRMATION]")

        while True:
            choice = input("\n>> Choose option [1/2/3] (Press Enter for 1): ").strip()
            if choice in ("", "1"):
                print("[SAFE MODE] Existing certificates will be preserved.")
                return "skip", planned_tasks
            elif choice == "2":
                print("[VERSIONING] New certificates will be saved with _v2 / _v3 suffixes.")
                return "version", planned_tasks
            elif choice == "3":
                confirm = input("[!] Are you SURE you want to overwrite existing certificates? Type 'yes' to confirm: ").strip().lower()
                if confirm == "yes":
                    print("[OVERWRITE] Overwrite confirmed. Existing certificates will be replaced.")
                    return "overwrite", planned_tasks
                else:
                    print("Overwrite not confirmed. Falling back to Safe [1] Skip existing.")
                    return "skip", planned_tasks
            else:
                print("Invalid choice. Please enter 1, 2, or 3.")

    def run_generation(
        self,
        planned_tasks: List[Tuple[Dict[str, Any], str, str, str]],
        mode: str
    ) -> Tuple[int, int, int, List[Dict[str, Any]]]:
        """
        Generates PDF and 300-DPI PNG certificates with live feedback.
        """
        print("\n[Step 4 of 4] Generating Certificates (Vector PDF + 300-DPI PNG)")
        print("-" * 68)

        successful = 0
        skipped_existing = 0
        failed = 0
        manifest_entries = []
        used_target_names = set()

        total = len(planned_tasks)

        for i, (rec, base_name, orig_pdf, orig_png) in enumerate(planned_tasks, 1):
            row_idx = rec.get("__row_index__", i)
            name = rec["name"]
            course = rec["course"]
            cert_id = rec["cert_id"]
            date_str = rec["date"]

            # Determine target filename according to overwrite safety mode
            if mode == "version":
                target_name = base_name
                v = 1
                while (target_name in used_target_names or
                       os.path.exists(os.path.join(self.output_dir, f"{target_name}.pdf")) or
                       os.path.exists(os.path.join(self.output_dir, f"{target_name}.png"))):
                    target_name = f"{base_name}_v{v}"
                    v += 1
                pdf_target = os.path.join(self.output_dir, f"{target_name}.pdf")
                png_target = os.path.join(self.output_dir, f"{target_name}.png")
            else:
                target_name = base_name
                dup_cnt = 1
                while target_name in used_target_names:
                    target_name = f"{base_name}_{dup_cnt}"
                    dup_cnt += 1
                pdf_target = os.path.join(self.output_dir, f"{target_name}.pdf")
                png_target = os.path.join(self.output_dir, f"{target_name}.png")

            used_target_names.add(target_name)

            # Skip handling
            if mode == "skip" and (os.path.exists(orig_pdf) and os.path.exists(orig_png)):
                print(f"  [{i}/{total}] [SKIP] Preserving existing: {name} ({target_name})")
                skipped_existing += 1
                manifest_entries.append({
                    "row": row_idx,
                    "name": name,
                    "course": course,
                    "cert_id": cert_id,
                    "date": date_str,
                    "status": "preserved_existing",
                    "files": {
                        "pdf": os.path.basename(orig_pdf),
                        "png": os.path.basename(orig_png)
                    }
                })
                continue

            print(f"  [{i}/{total}] Generating for: {name} ... ", end="", flush=True)

            try:
                # 1. Render Vector PDF
                self.renderer.render_certificate(rec, pdf_target)
                pdf_size_kb = os.path.getsize(pdf_target) // 1024

                # 2. Render 300-DPI PNG
                doc = pymupdf.open(pdf_target)
                page = doc[0]
                zoom = self.dpi / 72.0
                mat = pymupdf.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                pix.save(png_target)
                doc.close()
                png_size_kb = os.path.getsize(png_target) // 1024

                successful += 1
                print(f"[OK] Done! (PDF: {pdf_size_kb} KB, PNG: {png_size_kb} KB)")

                manifest_entries.append({
                    "row": row_idx,
                    "name": name,
                    "course": course,
                    "cert_id": cert_id,
                    "date": date_str,
                    "status": "generated",
                    "files": {
                        "pdf": os.path.basename(pdf_target),
                        "png": os.path.basename(png_target)
                    }
                })

            except Exception as e:
                failed += 1
                print(f"[FAILED] ERROR: {e}")
                self.logger.error(f"Row {row_idx}: Certificate generation failed for '{name}': {e}", exc_info=True)

        return successful, skipped_existing, failed, manifest_entries

    def save_manifest(
        self,
        total_records: int,
        successful: int,
        skipped_total: int,
        failed: int,
        manifest_entries: List[Dict[str, Any]],
        elapsed_seconds: float
    ):
        """Updates manifest.json with full run details."""
        manifest_path = os.path.join(self.output_dir, "manifest.json")
        try:
            with open(manifest_path, "w", encoding="utf-8") as f:
                json.dump({
                    "timestamp": datetime.now().isoformat(),
                    "total_records": total_records,
                    "successful": successful,
                    "skipped": skipped_total,
                    "failed": failed,
                    "elapsed_seconds": round(elapsed_seconds, 2),
                    "certificates": manifest_entries
                }, f, indent=2)
        except Exception as e:
            self.logger.warning(f"Could not write manifest.json: {e}")

    def display_final_summary(
        self,
        total_recipients: int,
        successful: int,
        skipped_invalid: int,
        skipped_existing: int,
        failed: int,
        elapsed: float,
        interactive: bool = True
    ):
        """Displays clear, user-friendly final summary (Requirement 10)."""
        total_skipped = skipped_invalid + skipped_existing
        abs_output = os.path.abspath(self.output_dir)

        print("\n" + "=" * 68)
        print("                     FINAL EXECUTION SUMMARY")
        print("=" * 68)
        print(f"  Total Recipients Ingested : {total_recipients}")
        print(f"  Successfully Generated    : {successful}")
        print(f"  Skipped (Total)           : {total_skipped}")
        if skipped_invalid > 0 or skipped_existing > 0:
            print(f"    - Incomplete/Invalid Rows : {skipped_invalid}")
            print(f"    - Preserved Existing Files: {skipped_existing}")
        print(f"  Failed (Errors)           : {failed}")
        print(f"  Processing Time           : {elapsed:.2f} seconds")
        print(f"  Output Folder             : {abs_output}")
        print("=" * 68)

        # Offer to open the folder in Windows File Explorer
        if interactive and hasattr(os, "startfile") and successful > 0:
            try:
                open_dir = input("\nWould you like to open the output folder in File Explorer? [Y/n]: ").strip().lower()
                if open_dir in ("", "y", "yes"):
                    print(f"Opening folder: {abs_output}")
                    os.startfile(abs_output)
            except Exception:
                pass


def main(cli_args: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(description="Certificate Generation Agent")
    parser.add_argument("-f", "--file", help="Recipient CSV or Excel file (skip file prompt)")
    parser.add_argument("--non-interactive", action="store_true", help="Run without interactive confirmation")
    parser.add_argument("--overwrite-mode", choices=["skip", "version", "overwrite"], default=None,
                        help="Collision resolution mode (skip, version, overwrite)")
    args = parser.parse_args(cli_args)

    agent = CertificateAgent()
    agent.print_banner()

    # Step 1: Recipient file selection
    if args.file:
        file_path = args.file
        if not os.path.exists(file_path):
            print(f"[!] Specified file not found: {file_path}")
            return
    else:
        file_path = agent.prompt_recipient_file()
        if not file_path:
            return

    # Step 2: Read & Validate
    valid_records, invalid_records = agent.inspect_and_validate(file_path)
    total_ingested = len(valid_records) + len(invalid_records)

    can_proceed = agent.display_validation_report(file_path, valid_records, invalid_records)
    if not can_proceed:
        return

    # User confirmation to proceed
    if not args.non_interactive:
        try:
            confirm = input(f"\n>> Generate certificates for {len(valid_records)} valid recipient(s)? [Y/n]: ").strip().lower()
            if confirm not in ("", "y", "yes"):
                print("Operation aborted by user.")
                return
        except (KeyboardInterrupt, EOFError):
            print("\nAborted.")
            return

    # Step 3: Overwrite safety check
    default_mode = args.overwrite_mode if args.overwrite_mode else ("skip" if args.non_interactive else None)
    mode, planned_tasks = agent.check_and_resolve_collisions(valid_records, default_choice=default_mode)

    # Step 4: Batch generation
    start_time = time.time()
    successful, skipped_existing, failed, manifest_entries = agent.run_generation(planned_tasks, mode)
    elapsed = time.time() - start_time

    # Save manifest
    agent.save_manifest(
        total_records=total_ingested,
        successful=successful,
        skipped_total=len(invalid_records) + skipped_existing,
        failed=failed,
        manifest_entries=manifest_entries,
        elapsed_seconds=elapsed
    )

    # Step 5: Final Summary
    agent.display_final_summary(
        total_recipients=total_ingested,
        successful=successful,
        skipped_invalid=len(invalid_records),
        skipped_existing=skipped_existing,
        failed=failed,
        elapsed=elapsed,
        interactive=not args.non_interactive
    )


if __name__ == "__main__":
    main()
