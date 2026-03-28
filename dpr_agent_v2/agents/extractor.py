"""
agents/extractor.py
═════════════════════
Shared async LLM call utility used by all handlers.
Uses httpx (bundled with anthropic SDK) — no aiohttp dependency.
"""

import json
import re
import asyncio
import httpx
from typing import Any


CLAUDE_API_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_MODEL  = "claude-sonnet-4-20250514"

EXTRACTION_SYSTEM = (
    "You are a financial data extraction assistant. "
    "Extract structured data from user messages and return ONLY valid JSON. "
    "No preamble, no explanation, no markdown fences. "
    "Use null for fields not mentioned. Numbers must be plain numbers without symbols."
)


def clean_json(raw: str) -> str:
    """Strip markdown fences and whitespace from LLM JSON output."""
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


async def llm_call(
    prompt: str,
    api_key: str,
    system: str = EXTRACTION_SYSTEM,
    model:  str = DEFAULT_MODEL,
    max_tokens: int = 1500,
) -> str:
    """Make a single async Claude API call. Returns raw text response."""
    headers = {
        "x-api-key":         api_key,
        "anthropic-version": "2023-06-01",
        "content-type":      "application/json",
    }
    payload = {
        "model":      model,
        "max_tokens": max_tokens,
        "system":     system,
        "messages":   [{"role": "user", "content": prompt}],
    }
    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(CLAUDE_API_URL, headers=headers, json=payload)
        data = resp.json()
        if resp.status_code != 200:
            raise RuntimeError(f"API error {resp.status_code}: {data}")
        return data["content"][0]["text"]


async def extract_json(
    prompt: str,
    api_key: str,
    fallback: dict,
    system: str = EXTRACTION_SYSTEM,
    model:  str = DEFAULT_MODEL,
    max_tokens: int = 1000,
) -> dict:
    """
    Call the LLM, parse JSON response, return fallback on any failure.
    Never raises — always returns a dict.
    """
    try:
        raw = await llm_call(prompt, api_key, system=system,
                             model=model, max_tokens=max_tokens)
        return json.loads(clean_json(raw))
    except Exception as e:
        print(f"[extractor] parse failed: {e}")
        return fallback


async def extract_all_parallel(
    prompts_and_fallbacks: list[tuple[str, dict]],
    api_key: str,
    model: str = DEFAULT_MODEL,
) -> list[dict]:
    """
    Run multiple extraction calls in parallel.
    Returns results in the same order as input.
    """
    tasks = [
        extract_json(prompt, api_key, fallback, model=model)
        for prompt, fallback in prompts_and_fallbacks
    ]
    return await asyncio.gather(*tasks)
