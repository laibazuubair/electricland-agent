import anthropic
import json
import os
from app.models import ExtractedDealData
from dotenv import load_dotenv

load_dotenv()

client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

def extract_deal_data(email_body: str) -> ExtractedDealData:

    prompt = f"""
You are a financial data extraction specialist for a UK land and property company.
Your job is to extract deal information from emails with MAXIMUM accuracy.

Extract ALL of the following fields from the email below.
If a field is not mentioned, set it to null.
If a value is ambiguous or unclear, set "confidence" to "low" for that field.

Return ONLY a valid JSON object with this exact structure — no explanation, no markdown, just the JSON:

{{
  "site_address": {{"value": "", "confidence": "high/medium/low"}},
  "site_area_acres": {{"value": null, "confidence": "high/medium/low"}},
  "purchase_price_gbp": {{"value": null, "confidence": "high/medium/low"}},
  "build_cost_gbp": {{"value": null, "confidence": "high/medium/low"}},
  "total_development_cost_gbp": {{"value": null, "confidence": "high/medium/low"}},
  "total_rent_roll_pa_gbp": {{"value": null, "confidence": "high/medium/low"}},
  "individual_units": {{
    "value": [
      {{"unit_name": "", "annual_rent_gbp": null}}
    ],
    "confidence": "high/medium/low"
  }},
  "target_yield_percent": {{"value": null, "confidence": "high/medium/low"}},
  "gdv_gbp": {{"value": null, "confidence": "high/medium/low"}},
  "profit_on_cost_percent": {{"value": null, "confidence": "high/medium/low"}},
  "development_timeline_months": {{"value": null, "confidence": "high/medium/low"}},
  "planning_status": {{"value": "", "confidence": "high/medium/low"}},
  "sender_name": {{"value": "", "confidence": "high/medium/low"}},
  "flags": []
}}

Important rules:
- All GBP values must be plain numbers e.g. 1850000 not £1,850,000
- If something looks inconsistent or contradictory, add a note to the "flags" array
- Be as accurate as possible — this data will be used for financial decisions

EMAIL TO EXTRACT FROM:
---
{email_body}
---
"""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()

    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip()

    data = json.loads(raw)
    return ExtractedDealData(**data)