"""
api_client/claude_client.py
────────────────────────────
Single wrapper for all Anthropic API calls in the DPR system.

Why this exists:
  - The API key is entered by the user in the Streamlit sidebar at runtime.
  - It must be injected into every call made by Orchestrator AND BenchmarkEngine.
  - Having one client means we change the key in one place, not six.

Usage:
    from api_client.claude_client import ClaudeClient

    client = ClaudeClient(api_key="sk-ant-...")
    response = await client.message(
        system="You are ...",
        user="Tell me about ...",
        max_tokens=1000
    )
    print(response)   # plain text string

    # Or for JSON extraction:
    data = await client.message_json(system=..., user=...)
"""

from __future__ import annotations
import json
import re
import asyncio
from typing import Optional


class ClaudeClient:
    """
    Thin wrapper around the Anthropic /v1/messages endpoint.
    Accepts an API key at construction time so Streamlit can pass
    the session-state key without any global state.
    """

    MODEL   = "claude-sonnet-4-20250514"
    API_URL = "https://api.anthropic.com/v1/messages"

    def __init__(self, api_key: str, model: str = None):
        if not api_key or not api_key.startswith("sk-ant-"):
            raise ValueError(
                "Invalid API key. It should start with 'sk-ant-'. "
                "Enter your key in the sidebar."
            )
        self.api_key = api_key
        self.model   = model or self.MODEL

    # ── Core async call ───────────────────────────────────────────────────────

    async def message(
        self,
        user: str,
        system: str = "",
        max_tokens: int = 1500,
        temperature: float = 0.3,
    ) -> str:
        """
        Send a message, return the text response as a plain string.
        Raises RuntimeError if the API call fails.
        """
        import aiohttp

        headers = {
            "Content-Type":      "application/json",
            "x-api-key":         self.api_key,
            "anthropic-version": "2023-06-01",
        }
        payload: dict = {
            "model":      self.model,
            "max_tokens": max_tokens,
            "messages":   [{"role": "user", "content": user}],
        }
        if system:
            payload["system"] = system

        async with aiohttp.ClientSession() as session:
            async with session.post(
                self.API_URL,
                json=payload,
                headers=headers,
                timeout=aiohttp.ClientTimeout(total=40),
            ) as resp:
                if resp.status != 200:
                    body = await resp.text()
                    raise RuntimeError(
                        f"Anthropic API error {resp.status}: {body[:300]}"
                    )
                data = await resp.json()
                # Extract text from content blocks
                return "".join(
                    block["text"]
                    for block in data.get("content", [])
                    if block.get("type") == "text"
                )

    async def message_json(
        self,
        user: str,
        system: str = "",
        max_tokens: int = 1500,
    ) -> dict:
        """
        Send a message expecting a JSON response.
        Strips code fences, parses, and returns the dict.
        Falls back to {} on parse failure.
        """
        raw = await self.message(user=user, system=system, max_tokens=max_tokens)
        return self._parse_json(raw)

    # ── Sync wrapper (for non-async Streamlit callbacks) ─────────────────────

    def message_sync(
        self,
        user: str,
        system: str = "",
        max_tokens: int = 1500,
    ) -> str:
        """Synchronous wrapper — safe to call from Streamlit event handlers."""
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # Already inside an event loop (e.g. Jupyter / some Streamlit versions)
                import concurrent.futures
                with concurrent.futures.ThreadPoolExecutor() as pool:
                    future = pool.submit(
                        asyncio.run,
                        self.message(user=user, system=system, max_tokens=max_tokens)
                    )
                    return future.result(timeout=45)
            else:
                return loop.run_until_complete(
                    self.message(user=user, system=system, max_tokens=max_tokens)
                )
        except Exception as e:
            raise RuntimeError(f"API call failed: {e}") from e

    def message_json_sync(
        self,
        user: str,
        system: str = "",
        max_tokens: int = 1500,
    ) -> dict:
        raw = self.message_sync(user=user, system=system, max_tokens=max_tokens)
        return self._parse_json(raw)

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _parse_json(text: str) -> dict:
        text = text.strip()
        text = re.sub(r'^```json\s*', '', text)
        text = re.sub(r'^```\s*',     '', text)
        text = re.sub(r'\s*```$',     '', text)
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            match = re.search(r'\{[\s\S]+\}', text)
            if match:
                try:
                    return json.loads(match.group())
                except Exception:
                    pass
        return {}

    @classmethod
    def validate_key(cls, key: str) -> tuple[bool, str]:
        """
        Returns (is_valid, error_message).
        Does a lightweight format check — does NOT make an API call.
        """
        if not key:
            return False, "API key is empty."
        key = key.strip()
        if not key.startswith("sk-ant-"):
            return False, "Key should start with 'sk-ant-'."
        if len(key) < 40:
            return False, "Key looks too short."
        return True, ""
