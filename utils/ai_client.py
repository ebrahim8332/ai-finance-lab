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
import streamlit as st

from utils.base import BaseProvider, FallbackTrigger, AllProvidersExhausted
from utils.groq_provider import GroqProvider, TIER1_MODEL, TIER2_MODEL, TIER3_MODEL, TIER4_MODEL, TIER5_MODEL

SESSION_LOCK_KEY = "locked_provider_index"


def build_chain() -> list[BaseProvider]:
    """
    Builds the ordered list of providers based on available API keys.

    Order when both keys are present:
      [0]  gemini-2.5-pro          (best quality)
      [1]  gemini-2.5-flash
      [2]  gemini-2.0-flash
      [3]  gemini-2.0-flash-lite
      [4]  gemini-2.5-flash-lite
      [5]  gemini-flash-latest
      [6]  llama-3.3-70b-versatile (Groq Tier 1)
      [7]  llama-4-scout            (Groq Tier 2)
      [8]  qwen3-32b                (Groq Tier 3)
      [9]  gpt-oss-120b             (Groq Tier 4)
      [10] llama-3.1-8b-instant     (Groq Tier 5 — last resort)
    """
    providers: list[BaseProvider] = [
        GroqProvider(TIER1_MODEL),
        GroqProvider(TIER2_MODEL),
        GroqProvider(TIER3_MODEL),
        GroqProvider(TIER4_MODEL),
        GroqProvider(TIER5_MODEL),
    ]

    if os.getenv("GEMINI_API_KEY"):
        from utils.gemini_provider import GeminiProvider
        for model in reversed([
            "gemini-2.5-pro",
            "gemini-2.5-flash",
            "gemini-2.0-flash",
            "gemini-2.0-flash-lite",
            "gemini-2.5-flash-lite",
            "gemini-flash-latest",
        ]):
            providers.insert(0, GeminiProvider(model))

    return providers


class FallbackChain:
    """
    Tries providers in order. Locks to the first one that succeeds.
    Re-locks if the locked model fails mid-session.
    """

    def __init__(self, providers: list[BaseProvider], session_state: dict):
        self.providers = providers
        self.session_state = session_state

    def complete(self, messages: list[dict], timeout: int = 90) -> tuple[str, str]:
        """Returns (response_text, model_name) from the first provider that succeeds."""
        start_index = self.session_state.get(SESSION_LOCK_KEY) or 0

        errors = []
        for i in range(start_index, len(self.providers)):
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
            "All AI models are currently unavailable. Please try again in a few minutes."
            + (f" Details: {details}" if details else "")
        )

    @property
    def locked_model(self) -> str | None:
        locked_index = self.session_state.get(SESSION_LOCK_KEY)
        if locked_index is not None and locked_index < len(self.providers):
            return self.providers[locked_index].model_name
        return None


# Cache the provider list so it is only built once per session
@st.cache_resource
def _build_cached_providers():
    return build_chain()


def get_chain(session_state: dict) -> FallbackChain:
    """Call this from any module to get a ready-to-use chain."""
    providers = _build_cached_providers()
    return FallbackChain(providers, session_state)
