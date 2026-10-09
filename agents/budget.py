"""Tracks cumulative cloud (Anthropic) API spend across runs and enforces
a hard cap. Once the cap is reached, the pipeline stops calling the cloud
API entirely and stays on the local model, even if local output fails
validation.

Spend is persisted to disk (not just in-memory) so the cap holds across
separate process/requests, not just within one run.

Adapted from ~/Desktop/mon-projet/src/bim_tender_check/budget.py (same
pattern, validated there against a real cloud escalation path).
"""

import json
import os
from pathlib import Path

STATE_PATH = Path(__file__).resolve().parents[1] / "data" / ".cloud_spend.json"

# USD/MTok, official rates from https://platform.claude.com/docs/en/about-claude/pricing.
# Update this table if you change CLOUD_MODEL to something not listed here --
# unknown models fail closed (spend can't be computed, so no cloud call is
# made) rather than silently going untracked.
PRICING_USD_PER_MTOK = {
    "claude-haiku-4-5-20251001": {"input": 1.0, "output": 5.0},
    "claude-sonnet-4-5": {"input": 3.0, "output": 15.0},
    "claude-sonnet-5": {"input": 2.0, "output": 10.0},
    "claude-opus-4-5": {"input": 5.0, "output": 25.0},
}


def _bare_model_name(model: str) -> str:
    return model.split("/", 1)[1] if "/" in model else model


def pricing_for(model: str) -> dict | None:
    return PRICING_USD_PER_MTOK.get(_bare_model_name(model))


def _usd_to_eur_rate() -> float:
    # Approximate, not a live FX rate -- override in .env if you want precision.
    return float(os.environ.get("USD_TO_EUR_RATE", "0.92"))


def cap_eur() -> float:
    return float(os.environ.get("CLOUD_BUDGET_EUR", "10"))


def _load_spent_usd() -> float:
    if STATE_PATH.exists():
        return json.loads(STATE_PATH.read_text()).get("spent_usd", 0.0)
    return 0.0


def _save_spent_usd(spent_usd: float) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    STATE_PATH.write_text(json.dumps({"spent_usd": spent_usd}, indent=2))


def spent_eur() -> float:
    return _load_spent_usd() * _usd_to_eur_rate()


def remaining_eur() -> float:
    return max(0.0, cap_eur() - spent_eur())


def has_budget() -> bool:
    return spent_eur() < cap_eur()


def estimate_cost_usd(model: str, prompt_tokens: int, completion_tokens: int, cached_prompt_tokens: int = 0) -> float | None:
    prices = pricing_for(model)
    if prices is None:
        return None
    return (
        prompt_tokens * prices["input"]
        + cached_prompt_tokens * prices["input"] * 0.1
        + completion_tokens * prices["output"]
    ) / 1_000_000


def record_spend_usd(usd: float) -> float:
    """Adds to the persisted running total. Returns the new total in EUR."""
    total_usd = _load_spent_usd() + usd
    _save_spent_usd(total_usd)
    return total_usd * _usd_to_eur_rate()
