import json
import os
from typing import List, Optional
from app.models import Deal

DB_PATH = "app/data/deals.json"

def _load() -> List[dict]:
    if not os.path.exists(DB_PATH):
        return []
    with open(DB_PATH, "r") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []

def _save(deals: List[dict]):
    with open(DB_PATH, "w") as f:
        json.dump(deals, f, indent=2)

def get_all_deals() -> List[dict]:
    return _load()

def get_deal_by_id(deal_id: str) -> Optional[dict]:
    deals = _load()
    for deal in deals:
        if deal["id"] == deal_id:
            return deal
    return None

def save_deal(deal: Deal) -> Deal:
    deals = _load()
    # Check if deal already exists and update it
    for i, d in enumerate(deals):
        if d["id"] == deal.id:
            deals[i] = deal.model_dump()
            _save(deals)
            return deal
    # Otherwise add new
    deals.append(deal.model_dump())
    _save(deals)
    return deal

def update_deal(deal_id: str, updates: dict) -> Optional[dict]:
    deals = _load()
    for i, deal in enumerate(deals):
        if deal["id"] == deal_id:
            deals[i].update(updates)
            _save(deals)
            return deals[i]
    return None

def delete_deal(deal_id: str) -> bool:
    deals = _load()
    new_deals = [d for d in deals if d["id"] != deal_id]
    if len(new_deals) == len(deals):
        return False
    _save(new_deals)
    return True