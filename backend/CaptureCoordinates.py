import argparse
import os
import sys
from typing import Dict, List, Optional

import fitz
import pandas as pd


def classify_text(text: str) -> Optional[str]:
    """Classify the extracted text into Instrument, Pipe-run, or Equipment."""
    cleaned = text.strip()
    if not cleaned:
        return None

    has_dash = "-" in cleaned
    has_quote = '"' in cleaned
    is_alnum = cleaned.isalnum()

    if has_quote and has_dash:
        return "Pipe-run"
    if has_dash and not has_quote:
        return "Equipment"
    if is_alnum and not has_dash and not has_quote:
        return "Instrument"
    return None


def extract_text_objects_from_pdf(pdf_path: str, pid_no: str, starting_id: int = 1) -> List[Dict[str, object]]:
    """Extract text objects and bounding boxes from a PDF using PyMuPDF."""
    rows: List[Dict[str, object]] = []
    box_id = starting_id

    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        print(f"Warning: Failed to open PDF '{pdf_path}': {exc}", file=sys.stderr)
        return rows

    for page_index in range(len(doc)):
        page = doc[page_index]
        page_no = page_index + 1
        words = page.get_text("words")
        words.sort(key=lambda item: (item[1], item[0]))

        for x1, y1, x2, y2, word, block_no, line_no, word_no in words:
            text = str(word).strip()
            if not text:
                continue

            tag_type = classify_text(text)
            if not tag_type:
                continue

            rows.append(
                {
                    "P&ID no.": pid_no,
                    "Text": text,
                    "Box ID": box_id,
                    "Tag Type": tag_type,
                    "X1": x1,
                    "Y1": y1,
                    "X2": x2,
                    "Y2": y2,
                    "Page No.": page_no,
                    
                }
            )
            box_id += 1

    doc.close()
    return rows


def process_folder(input_folder: str, output_file: str) -> None:
    """Process all PDFs in the input folder and export results to Excel."""
    if not os.path.isdir(input_folder):
        raise FileNotFoundError(f"Input folder not found: {input_folder}")

    pdf_files = [
        f for f in sorted(os.listdir(input_folder)) if f.lower().endswith(".pdf")
    ]
    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in folder: {input_folder}")

    output_rows: List[Dict[str, object]] = []
    next_box_id = 1

    for pdf_name in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_name)
        pid_no = os.path.splitext(pdf_name)[0]
        pdf_rows = extract_text_objects_from_pdf(pdf_path, pid_no, starting_id=next_box_id)
        output_rows.extend(pdf_rows)
        next_box_id += len(pdf_rows)

    df = pd.DataFrame(output_rows, columns=[
        "P&ID no.",
        "Text",
        "Tag Type",
        "Box ID",
        "X1",
        "Y1",
        "X2",
        "Y2",
        "Page No.",      
        
    ])

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="capture_coordinates")

    print(f"Exported {len(output_rows)} rows to {output_file}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract text objects from PDFs and export to Excel."
    )
    parser.add_argument(
        "input_folder",
        help="Path to the folder containing PDF files.",
    )
    parser.add_argument(
        "output_file",
        nargs="?",
        default="capture_coordinates.xlsx",
        help="Excel output file path (default: capture_coordinates.xlsx).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        process_folder(args.input_folder, args.output_file)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
