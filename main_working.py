# main_working.py
"""
Main Script: Orchestrates PDF extraction, matching, and Excel update.
Run this: python main_working.py
"""

import os
import logging
from document_extractor import DocumentExtractor
from simple_fuzzy_matcher import SimpleFuzzyMatcher
# print(SimpleFuzzyMatcher.__file__)
from excel_updater import ExcelUpdater
from config import (
    EXCEL_TEMPLATE_PATH, OUTPUT_EXCEL_PATH, FUZZY_MATCH_THRESHOLD
)

logging.basicConfig(level=logging.INFO)

# List your PDF files here (scalable: add more paths)


PDF_PATHS = [
    r"C:\Users\User\Desktop\QBS\indus_motor_demo\RFQ-1.pdf",
    r"C:\Users\User\Desktop\QBS\indus_motor_demo\RFQ-2.pdf",
    r"C:\Users\User\Desktop\QBS\indus_motor_demo\RFQ-3.pdf",
    # Add more: "RFQ-4.pdf", etc.
]

def main():
    # Step 1: Extract items from all PDFs using LLM
    extractor = DocumentExtractor()
    all_items = []
    for pdf_path in PDF_PATHS:
        if not os.path.exists(pdf_path):
            logging.error(f"PDF not found: {pdf_path}")
            continue
        items = extractor.extract_items_from_pdf(pdf_path)
        all_items.append(items)
        logging.info(f"Extracted {len(items)} items from {pdf_path}")

    # Step 2: Match items across PDFs
    matcher = SimpleFuzzyMatcher(threshold=FUZZY_MATCH_THRESHOLD)
    matched_data = matcher.match_items(all_items)
    logging.info(f"Matched {len(matched_data)} unique items")

    # Step 3: Update Excel
    updater = ExcelUpdater(EXCEL_TEMPLATE_PATH)
    updater.update_comparative(matched_data)
    os.makedirs(os.path.dirname(OUTPUT_EXCEL_PATH), exist_ok=True)
    updater.save(OUTPUT_EXCEL_PATH)
    logging.info(f"Updated Excel saved to {OUTPUT_EXCEL_PATH}")

if __name__ == "__main__":
    main()