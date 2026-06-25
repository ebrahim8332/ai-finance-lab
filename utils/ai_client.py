"""
AI client for FinanceAI Lab.

Builds a fallback chain: Gemini models first (best quality, free tier),
then Groq models as backup. On the first successful call, the chain locks
to that model for the rest of the session. If the locked model fails later
(e.g. rate limit), the chain continues from that point and re-locks.

Usage in any module:
    from utils.ai_client import get_chain
    chain = get_chain(st.session_state)
    response, model_used = chain.complete(messages)
"""

import os

from utils.base import BaseProvider, FallbackTrigger, AllProvidersExhausted
from utils.groq_provider import GroqProvider, TIER1_MODEL, TIER2_MODEL, TIER3_MODEL, TIER4_MODEL, TIER5_MODEL, TIER6_MODEL

SESSION_LOCK_KEY = "locked_provider_index"


def build_chain() -> list[BaseProvider]:
    """
    Builds the ordered list of providers based on available API keys.

    Order when both keys are present:
      [0]  gemini-2.5-pro
      [1]  gemini-3-flash-preview
      [2]  gemini-3.1-flash-lite
      [3]  gemini-2.5-flash
      [4]  gemini-2.5-flash-lite
      [5]  gemini-2.0-flash            (deprecated June 2026, 8K output cap)
      [6]  gemini-2.0-flash-lite       (deprecated June 2026, 8K output cap)
      [7]  llama-3.3-70b-versatile     (Groq Tier 1)
      [8]  qwen3.6-27b              (Groq Tier 2)
      [9]  qwen3-32b                   (Groq Tier 3)
      [10] gpt-oss-120b                (Groq Tier 4)
      [11] llama-3.1-8b-instant        (Groq Tier 5)
      [12] gpt-oss-20b                 (Groq Tier 6)
      [13] gemini-flash-latest         (unstable alias — absolute last resort)
    """
    providers: list[BaseProvider] = [
        GroqProvider(TIER1_MODEL),
        GroqProvider(TIER2_MODEL),
        GroqProvider(TIER3_MODEL),
        GroqProvider(TIER4_MODEL),
        GroqProvider(TIER5_MODEL),
        GroqProvider(TIER6_MODEL),
    ]

    if os.getenv("GEMINI_API_KEY"):
        from utils.gemini_provider import GeminiProvider
        for model in reversed([
            "gemini-2.5-pro",
            "gemini-3-flash-preview",
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash",
            "gemini-2.5-flash-lite",
            "gemini-2.0-flash",
            "gemini-2.0-flash-lite",
        ]):
            providers.insert(0, GeminiProvider(model))
        # gemini-flash-latest is an unstable alias — appended last, after all Groq models
        providers.append(GeminiProvider("gemini-flash-latest"))

    return providers


class FallbackChain:
    """
    Tries each provider in order. Locks to the first one that succeeds.

    If the locked provider fails on a later call (e.g. mid-session rate limit),
    the chain continues from that point and re-locks to the next provider that
    succeeds. Only raises AllProvidersExhausted when every remaining provider fails.
    """

    def __init__(self, providers: list[BaseProvider], session_state: dict):
        self.providers = providers
        self.session_state = session_state

    def complete(self, messages: list[dict], timeout: int = 90) -> tuple[str, str]:
        """Returns (response_text, model_name) from the first provider that succeeds."""
        errors = []
        for i in range(len(self.providers)):
            provider = self.providers[i]
            try:
                response = provider.complete(messages, timeout=timeout)
                self.session_state[SESSION_LOCK_KEY] = i
                return response, provider.model_name
            except FallbackTrigger as e:
                errors.append(f"{provider.model_name}: {e}")
                self.session_state["_fallback_errors"] = list(errors)
                continue

        details = " | ".join(errors[-3:])
        raise AllProvidersExhausted(
            "All available models are currently rate-limited or unavailable. "
            "Please try again in a few minutes."
            + (f" Last attempts: {details}" if details else "")
        )

    @property
    def locked_model(self) -> str | None:
        locked_index = self.session_state.get(SESSION_LOCK_KEY)
        if locked_index is not None and locked_index < len(self.providers):
            return self.providers[locked_index].model_name
        return None


def get_chain(session_state: dict) -> FallbackChain:
    """Call this from any module to get a ready-to-use chain."""
    providers = build_chain()
    return FallbackChain(providers, session_state)
