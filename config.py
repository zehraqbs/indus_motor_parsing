# config.py
"""
Configuration settings for the PDF to Excel processor.
"""

# Excel configuration
EXCEL_TEMPLATE_PATH = r"C:\Users\User\Desktop\QBS\indus_motor_demo\Local Comparative.xlsx"  # Path to your Excel template
OUTPUT_EXCEL_PATH = r"C:\Users\User\Desktop\QBS\indus_motor_demo\output\Updated_Comparative.xlsx"  # Where to save the updated file

# Columns in Excel (adjust if your template changes)
EXCEL_COLUMNS = {
    'description': 'C',  # Where to put item descriptions
    'quantity': 'D',     # Where to put quantities
}

# Vendor price columns (maps vendor labels A, B, C... to Excel columns)
# Extend this dictionary for more vendors (e.g., 'D': 'P', 'E': 'R')
PURCHASE_HISTORY_COLUMNS = {
    'A': 'I',  # Vendor A prices
    'B': 'L',  # Vendor B prices
    'C': 'N',  # Vendor C prices
}

# Row where data starts in the Excel sheet
DATA_START_ROW = 7  # Adjust based on your template (e.g., after headers)

# Fuzzy matching threshold (0-100, higher = stricter matching)
FUZZY_MATCH_THRESHOLD = 75

# LLM configuration (for extraction)
GROQ_MODEL = "llama-3.1-70b-versatile"  # Or other models like "mixtral-8x7b-32768"
EXTRACTION_PROMPT_TEMPLATE = """
Extract all items from this RFQ/Quotation document text. 
Focus on the table containing items, descriptions, quantities, and prices.

For each item, return a dictionary with:
- 'description': str (cleaned item name/description)
- 'quantity': float (quantity required/quoted, e.g., 4.0)
- 'uom': str (unit of measure, e.g., 'EA', 'SET')
- 'unit_price': float or None (unit price if provided, else None)

Return ONLY a JSON list of these dictionaries. No other text.

Document text:
{text}
"""