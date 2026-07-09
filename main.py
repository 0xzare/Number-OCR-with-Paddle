#!/usr/bin/env python3
"""Number OCR with PaddleOCR.

Extracts phone numbers (10 digits) and SIM serial numbers (15 digits) from
images in a folder and saves the results to an Excel file.

Usage:
    python main.py [folder] [-o output.xlsx] [--device gpu:0] [--lang en] [-r]

Run ``python main.py -h`` for all options. Can also be used as a library:
    from main import process_images_from_folder
    process_images_from_folder("path/to/images", "results.xlsx")
"""

from __future__ import annotations

import argparse
import logging
import re
import time
from datetime import datetime
from functools import lru_cache
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter

logger = logging.getLogger("number_ocr")

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}
PHONE_DIGITS = 10
SERIAL_DIGITS = 15
MAX_COLUMN_WIDTH = 50
DEFAULT_DEVICE = "gpu:0"
DEFAULT_LANG = "en"

_NON_DIGITS = re.compile(r"\D+")


@lru_cache(maxsize=1)
def get_ocr(lang: str = DEFAULT_LANG, device: str = DEFAULT_DEVICE):
    """Create the PaddleOCR engine once, on first use.

    The import happens lazily because importing paddle is slow; this keeps
    ``import main`` cheap and lets ``--help`` respond instantly.
    """
    from paddleocr import PaddleOCR

    return PaddleOCR(lang=lang, device=device, use_angle_cls=True)


def find_numeric_elements(texts) -> tuple[str | None, str | None]:
    """Return the first 10-digit (phone) and 15-digit (serial) numbers found.

    Non-digit separators (spaces, dashes, dots) inside a detected text are
    ignored, so "0912 345 6789" is still recognized as a phone number.
    """
    phone = serial = None
    for text in texts:
        digits = _NON_DIGITS.sub("", str(text))
        if phone is None and len(digits) == PHONE_DIGITS:
            phone = digits
        elif serial is None and len(digits) == SERIAL_DIGITS:
            serial = digits
        if phone is not None and serial is not None:
            break
    return phone, serial


def process_single_image(
    image_path: str | Path,
    lang: str = DEFAULT_LANG,
    device: str = DEFAULT_DEVICE,
) -> tuple[str | None, str | None, str]:
    """Run OCR on a single image and extract the phone and serial numbers."""
    try:
        ocr_result = get_ocr(lang, device).ocr(str(image_path), cls=True)
        if not ocr_result or not ocr_result[0]:
            return None, None, "No text detected"
        texts = [text for _box, (text, _confidence) in ocr_result[0]]
        phone, serial = find_numeric_elements(texts)
        return phone, serial, "Success"
    except Exception as exc:  # one bad image must not stop the whole batch
        logger.error("Error processing %s: %s", image_path, exc)
        return None, None, f"Error: {exc}"


def find_images(folder: str | Path, recursive: bool = False) -> list[Path]:
    """Return supported image files (case-insensitive extensions), sorted."""
    folder = Path(folder)
    iterator = folder.rglob("*") if recursive else folder.iterdir()
    return sorted(
        path
        for path in iterator
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def save_to_excel(df: pd.DataFrame, output_path: str | Path) -> None:
    """Write the results DataFrame to Excel with auto-fitted column widths."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="OCR Results", index=False)
        worksheet = writer.sheets["OCR Results"]
        for index, column in enumerate(df.columns, start=1):
            content_width = int(df[column].astype(str).str.len().max() or 0)
            width = min(max(content_width, len(column)) + 2, MAX_COLUMN_WIDTH)
            worksheet.column_dimensions[get_column_letter(index)].width = width


def process_images_from_folder(
    folder_path: str | Path,
    output_excel_path: str | Path | None = None,
    lang: str = DEFAULT_LANG,
    device: str = DEFAULT_DEVICE,
    recursive: bool = False,
) -> pd.DataFrame | None:
    """Process all images in a folder and save the results to Excel."""
    image_files = find_images(folder_path, recursive=recursive)
    if not image_files:
        logger.warning("No image files found in the folder!")
        return None

    logger.info("Found %d images.", len(image_files))

    results = []
    for i, image_path in enumerate(image_files, 1):
        logger.info("Processing image %d/%d: %s", i, len(image_files), image_path.name)

        phone, serial, status = process_single_image(image_path, lang, device)
        results.append(
            {
                "Filename": image_path.name,
                "Full Path": str(image_path.resolve()),
                "Phone Number": phone or "",
                "Serial Number": serial or "",
                "Status": status,
                "Processing Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        )

        logger.info("  - Phone Number: %s", phone or "Not found")
        logger.info("  - Serial Number: %s", serial or "Not found")
        logger.info("  - Status: %s", status)
        logger.info("-" * 50)

    df = pd.DataFrame(results)

    if output_excel_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel_path = f"ocr_results_{timestamp}.xlsx"

    try:
        save_to_excel(df, output_excel_path)
    except Exception as exc:
        logger.error("Error saving Excel file: %s", exc)
        return df

    logger.info("\nResults successfully saved to file: %s", output_excel_path)
    logger.info("\nSummary Statistics:")
    logger.info("- Total images: %d", len(image_files))
    logger.info("- Successfully processed: %d", (df["Status"] == "Success").sum())
    logger.info("- Phone numbers found: %d", (df["Phone Number"] != "").sum())
    logger.info("- Serial numbers found: %d", (df["Serial Number"] != "").sum())
    return df


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract phone and serial numbers from SIM-card images."
    )
    parser.add_argument(
        "folder",
        nargs="?",
        default=".",
        help="Folder containing the images (default: current folder)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output Excel file path (default: ocr_results_<timestamp>.xlsx)",
    )
    parser.add_argument(
        "--device",
        default=DEFAULT_DEVICE,
        help=f'Inference device, e.g. "gpu:0" or "cpu" (default: {DEFAULT_DEVICE})',
    )
    parser.add_argument(
        "--lang",
        default=DEFAULT_LANG,
        help=f"OCR language (default: {DEFAULT_LANG})",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="Also search subfolders for images",
    )
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    args = parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        logger.error("The specified folder does not exist: %s", folder)
        return 1

    logger.info("\nStarting image processing from folder: %s", folder.resolve())
    start = time.perf_counter()
    process_images_from_folder(
        folder,
        args.output,
        lang=args.lang,
        device=args.device,
        recursive=args.recursive,
    )
    logger.info("Calculation Time: %.2f s", time.perf_counter() - start)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
