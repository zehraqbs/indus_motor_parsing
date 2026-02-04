# document_extractor.py   ── LLM-powered version ──

import os
import json
from groq import Groq
from pypdf import PdfReader
from dotenv import load_dotenv

load_dotenv()

class DocumentExtractor:
    def __init__(self):
        print("GROQ_API_KEY =", os.getenv("GROQ_API_KEY"))
        self.client = Groq(api_key=os.getenv("GROQ_API_KEY"))
        self.model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        try:
            reader = PdfReader(pdf_path)
            text = ""
            for page in reader.pages:
                page_text = page.extract_text() or ""
                text += page_text + "\n"
            return text.strip()
        except Exception as e:
            raise ValueError(f"Failed to read PDF {pdf_path}: {e}")

    def extract_items_from_pdf(self, pdf_path: str) -> list[dict]:
        text = self.extract_text_from_pdf(pdf_path)

        prompt = f"""Extract structured item data from this quotation / RFQ document.
Return ONLY a valid JSON array of objects. No explanation, no markdown, no extra text.

Each object must have exactly these keys:
- "description": str   (clean item name / description)
- "quantity": float    (use the number shown — e.g. 4, 65, 200)
- "uom": str           ("EA", "SET", etc. — use "EA" if unclear)
- "unit_price": float or null   (the price per unit if shown, else null)

Document text:
{text}
"""

        try:
            completion = self.client.chat.completions.create(
                messages=[{"role": "user", "content": prompt}],
                model=self.model,
                temperature=0.0,
                max_tokens=1500,
            )

            raw = completion.choices[0].message.content.strip()

            # Try to clean common LLM wrappers
            if raw.startswith("```json"):
                raw = raw.split("```json", 1)[1].split("```", 1)[0].strip()
            elif raw.startswith("```"):
                raw = raw.split("```", 2)[1].strip()

            items = json.loads(raw)

            if not isinstance(items, list):
                raise ValueError("LLM did not return a list")

            # Basic normalization
            for item in items:
                if item.get("quantity") is not None:
                    try:
                        item["quantity"] = float(item["quantity"])
                    except:
                        item["quantity"] = None
                if item.get("unit_price") is not None:
                    try:
                        item["unit_price"] = float(item["unit_price"])
                    except:
                        item["unit_price"] = None

            return items

        except json.JSONDecodeError as e:
            print("JSON parse failed. Raw LLM output:")
            print(raw)
            raise
        except Exception as e:
            raise RuntimeError(f"LLM extraction failed for {pdf_path}: {e}")