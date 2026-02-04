# simple_fuzzy_matcher.py
"""
Simple Fuzzy Matcher: Matches item descriptions across different PDFs.
Uses fuzzywuzzy for similarity scoring.
"""

from fuzzywuzzy import fuzz
from typing import List, Dict
import logging

logging.basicConfig(level=logging.INFO)

class SimpleFuzzyMatcher:
    def __init__(self, threshold: int = 75):
        """
        Args:
            threshold: Matching score threshold (0-100)
        """
        self.threshold = threshold

    def match_items(self, all_items: List[List[Dict]]) -> Dict[str, Dict]:
        """
        Match items across multiple PDF extractions using fuzzy matching.
        
        Args:
            all_items: List of lists (one per PDF) of item dicts.
        
        Returns:
            Dict of matched groups: {canonical_description: {'quantity': float, 'prices': [float|None, ...]}}
        """
        if not all_items:
            return {}

        # Use the first PDF as the base (assuming it's the main RFQ with all items)
        base_items = all_items[0]
        matched_groups = {}

        for base_item in base_items:
            desc = base_item['description'].strip().lower()
            group = {
                'quantity': base_item['quantity'],
                'uom': base_item['uom'],
                'prices': [base_item['unit_price']] + [None] * (len(all_items) - 1)
            }

            # Match against other PDFs
            for pdf_idx, pdf_items in enumerate(all_items[1:], start=1):
                best_match = None
                best_score = 0
                for item in pdf_items:
                    score = fuzz.token_sort_ratio(desc, item['description'].strip().lower())
                    if score > best_score and score >= self.threshold:
                        best_score = score
                        best_match = item

                if best_match:
                    group['prices'][pdf_idx] = best_match['unit_price']
                    logging.info(f"Matched '{desc}' to '{best_match['description']}' (score: {best_score})")

            # Use the base description as canonical
            matched_groups[base_item['description']] = group

        # Report any unmatched items (for debugging)
        self._log_unmatched(all_items)

        return matched_groups

    def _log_unmatched(self, all_items: List[List[Dict]]):
        for pdf_idx, pdf_items in enumerate(all_items[1:], start=1):
            for item in pdf_items:
                matched = False
                for base_item in all_items[0]:
                    score = fuzz.token_sort_ratio(
                        base_item['description'].strip().lower(), 
                        item['description'].strip().lower()
                    )
                    if score >= self.threshold:
                        matched = True
                        break
                if not matched:
                    logging.warning(f"Unmatched item in PDF {pdf_idx + 1}: {item['description']}")