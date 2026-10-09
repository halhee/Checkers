"""Local-first, cloud-on-demand LLM execution. Each task runs on the local
Ollama model first; only if that output is missing or fails its grounding
check does it get re-run on the cloud model -- and only if there's still
budget left under the persisted spend cap (see budget.py).

Adapted from ~/Desktop/mon-projet/src/bim_tender_check/llm_tiers.py.
"""

import os

from crewai import LLM, Crew, Process

from . import budget

DEFAULT_CLOUD_MODEL = "anthropic/claude-haiku-4-5-20251001"


def build_local_llm() -> LLM:
    return LLM(
        model=os.environ.get("OLLAMA_MODEL", "ollama/llama3.2:3b"),
        base_url=os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434"),
    )


def _cloud_model_name() -> str:
    return os.environ.get("CLOUD_MODEL", DEFAULT_CLOUD_MODEL)


def build_cloud_llm() -> LLM | None:
    api_key = os.environ.get("CLOUD_ANTHROPIC_API_KEY")
    if not api_key:
        return None

    model = _cloud_model_name()
    if budget.pricing_for(model) is None:
        print(f"  [{model} isn't in budget.PRICING_USD_PER_MTOK -- can't guarantee the spend cap, staying local]")
        return None

    if not budget.has_budget():
        print(f"  [cloud budget of €{budget.cap_eur():.2f} reached (€{budget.spent_eur():.2f} spent so far) -- staying local]")
        return None

    return LLM(
        model=model,
        api_key=api_key,
        # This host may have ANTHROPIC_BASE_URL set for Claude Code itself --
        # pin the real Anthropic endpoint explicitly so this call can't inherit it.
        base_url="https://api.anthropic.com",
    )


def _record_cloud_spend(crew: Crew, model: str) -> None:
    usage = crew.usage_metrics
    if usage is None:
        print("  [no usage metrics returned -- could not record cloud spend]")
        return
    cost_usd = budget.estimate_cost_usd(
        model,
        prompt_tokens=usage.prompt_tokens or 0,
        completion_tokens=usage.completion_tokens or 0,
        cached_prompt_tokens=usage.cached_prompt_tokens or 0,
    )
    if cost_usd is None:
        return
    new_total_eur = budget.record_spend_usd(cost_usd)
    print(f"  [cloud call cost ~${cost_usd:.4f} -- cumulative spend now ~€{new_total_eur:.2f} of €{budget.cap_eur():.2f} cap]")


def run_with_fallback(build_agent_and_task, validator) -> tuple[str, str]:
    """`build_agent_and_task(llm)` returns (agent, task) built against that
    LLM. Runs it on the local model first; falls back to cloud (if budget
    allows) when the local run errors or fails `validator(text) -> bool`.
    Returns (output_text, tier_used)."""
    local_llm = build_local_llm()
    agent, task = build_agent_and_task(local_llm)
    text = None
    try:
        crew = Crew(agents=[agent], tasks=[task], process=Process.sequential, verbose=False)
        text = str(crew.kickoff())
        if validator(text):
            return text, "local"
    except Exception as exc:
        print(f"  [local run failed: {exc}]")

    cloud_llm = build_cloud_llm()  # returns None if no key, unpriced model, or budget exhausted
    if cloud_llm is None:
        if text is not None:
            print("  [local output failed validation and cloud fallback unavailable -- keeping local result]")
            return text, "local (unvalidated)"
        raise RuntimeError(
            "Local model failed and cloud fallback is unavailable (no CLOUD_ANTHROPIC_API_KEY, "
            "unpriced model, or budget cap reached). Get Ollama working, or check .env, to proceed."
        )

    agent, task = build_agent_and_task(cloud_llm)
    crew = Crew(agents=[agent], tasks=[task], process=Process.sequential, verbose=False)
    cloud_text = str(crew.kickoff())
    _record_cloud_spend(crew, _cloud_model_name())
    if not validator(cloud_text):
        print("  [cloud output also failed grounding validation -- reporting it anyway, flag for review]")
    return cloud_text, "cloud"
