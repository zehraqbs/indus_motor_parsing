# excel_updater.py
"""
Excel Updater: Updates specific cells in Excel without touching formulas.
Uses openpyxl for formula preservation.
"""

import openpyxl
from typing import Dict
from config import EXCEL_COLUMNS, PURCHASE_HISTORY_COLUMNS, DATA_START_ROW

class ExcelUpdater:
    def __init__(self, template_path: str):
        self.wb = openpyxl.load_workbook(template_path)
        self.sheet = self.wb.active  # Assuming single sheet; adjust if needed

    def update_comparative(self, matched_data: Dict[str, Dict]):
        """
        Update Excel with matched data.
        
        Args:
            matched_data: From SimpleFuzzyMatcher
        """
        row = DATA_START_ROW
        
        vendor_keys = sorted(PURCHASE_HISTORY_COLUMNS.keys())  # A, B, C...
        
        for desc, data in matched_data.items():
            # Write description
            desc_col = EXCEL_COLUMNS['description']
            self.sheet[f"{desc_col}{row}"] = desc
            
            # Write quantity
            qty_col = EXCEL_COLUMNS['quantity']
            self.sheet[f"{qty_col}{row}"] = data['quantity']
            
            # Write prices for each vendor
            prices = data['prices']
            for vendor_idx, vendor_key in enumerate(vendor_keys):
                if vendor_idx < len(prices):
                    price_col = PURCHASE_HISTORY_COLUMNS[vendor_key]
                    self.sheet[f"{price_col}{row}"] = prices[vendor_idx] if prices[vendor_idx] is not None else ""
            
            row += 1

    def save(self, output_path: str):
        self.wb.save(output_path)