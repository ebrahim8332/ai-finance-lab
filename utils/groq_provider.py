import os
import groq as groq_errors
from groq import Groq

from utils.base import BaseProvider, FallbackTrigger

# Groq models used in the fallback chain, ordered by preference.
TIER1_MODEL = "llama-3.3-70b-versatile"
TIER2_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"
TIER3_MODEL = "qwen/qwen3-32b"
TIER4_MODEL = "openai/gpt-oss-120b"
TIER5_MODEL = "llama-3.1-8b-instant"
TIER6_MODEL = "openai/gpt-oss-20b"                     # smaller/faster sibling to 120B, last Groq resort


class GroqProvider(BaseProvider):
    """
    Calls the Groq API. One instance per model — pass the model name at construction.
    """

    def __init__(self, model_name: str):
        self.model_name = model_name
        self.client = Groq(api_key=os.getenv("GROQ_API_KEY"))

    def complete(self, messages: list[dict], timeout: int = 60, temperature: float = 0.3) -> str:
        """
        Sends messages to Groq and returns the response text.
        Raises FallbackTrigger on rate limit, unavailability, timeout, or connection error.
        Auth errors and unexpected errors are re-raised as-is.
        """
        try:
            return self._call(messages, timeout, temperature)

        except groq_errors.RateLimitError as e:
            raise FallbackTrigger(
                f"Groq rate limit on {self.model_name}"
            ) from e

        except groq_errors.APIStatusError as e:
            if e.status_code == 503:
                raise FallbackTrigger(
                    f"Groq model unavailable: {self.model_name}"
                ) from e
            if e.status_code == 413:
                raise FallbackTrigger(
                    f"Groq request too large for {self.model_name} — falling back"
                ) from e
            raise  # 401 auth errors and others should surface, not fall back

        except groq_errors.APITimeoutError as e:
            raise FallbackTrigger(
                f"Groq timeout on {self.model_name}"
            ) from e

        except groq_errors.APIConnectionError as e:
            # One retry on connection error before triggering fallback
            try:
                return self._call(messages, timeout, temperature)
            except Exception as retry_error:
                raise FallbackTrigger(
                    f"Groq connection error on {self.model_name} (failed after retry)"
                ) from retry_error

    def _call(self, messages: list[dict], timeout: int, temperature: float = 0.3) -> str:
        max_tokens = int(os.getenv("GROQ_MAX_COMPLETION_TOKENS", "4000"))
        response = self.client.chat.completions.create(
            model=self.model_name,
            messages=messages,
            max_tokens=max_tokens,
            temperature=temperature,
            timeout=timeout,
        )
        return response.choices[0].message.content
