import os
import groq as groq_errors
from groq import Groq

from utils.base import BaseProvider, FallbackTrigger

TIER1_MODEL = "llama-3.3-70b-versatile"
TIER2_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"
TIER3_MODEL = "qwen/qwen3-32b"
TIER4_MODEL = "openai/gpt-oss-120b"
TIER5_MODEL = "llama-3.1-8b-instant"


class GroqProvider(BaseProvider):
    """Calls the Groq API. One instance per model."""

    def __init__(self, model_name: str):
        self.model_name = model_name
        self.client = Groq(api_key=os.getenv("GROQ_API_KEY"))

    def complete(self, messages: list[dict], timeout: int = 60, temperature: float = 0.3) -> str:
        try:
            return self._call(messages, timeout, temperature)

        except groq_errors.RateLimitError as e:
            raise FallbackTrigger(f"Groq rate limit on {self.model_name}") from e

        except groq_errors.APIStatusError as e:
            if e.status_code in (503, 413):
                raise FallbackTrigger(f"Groq unavailable/too large: {self.model_name}") from e
            raise

        except groq_errors.APITimeoutError as e:
            raise FallbackTrigger(f"Groq timeout on {self.model_name}") from e

        except groq_errors.APIConnectionError as e:
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
