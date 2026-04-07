"""LLM client wrapper supporting Anthropic and Google Gemini."""

import json
import logging
import re
from typing import Any

from app.config import settings

logger = logging.getLogger(__name__)


class LLMClient:
    """Unified wrapper for LLM API calls with retry and JSON parsing."""

    def __init__(
        self,
        provider: str | None = None,
        api_key: str | None = None,
        model: str | None = None,
    ):
        self._provider = provider or settings.llm_provider
        self._api_key = api_key
        self._model = model
        self._client = None

        if self._provider == "anthropic":
            self._api_key = self._api_key or settings.anthropic_api_key
            self._model = self._model or settings.llm_model
        elif self._provider == "google":
            self._api_key = self._api_key or settings.google_api_key
            self._model = self._model or settings.google_model

    async def generate_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str | None = None,
        max_tokens: int = 8192,
    ) -> dict[str, Any]:
        """Call LLM and parse response as JSON."""
        last_error = None

        for attempt in range(settings.max_retries):
            try:
                text = await self._call_llm(
                    system_prompt, user_prompt, model, max_tokens
                )
                text = _extract_json_block(text)
                return json.loads(text)

            except json.JSONDecodeError as e:
                logger.warning("JSON parse error on attempt %d: %s", attempt + 1, str(e))
                last_error = e
                user_prompt += "\n\n重要：請務必只回傳純 JSON 格式，不要包含任何其他文字。"

            except Exception as e:
                logger.warning("LLM error on attempt %d: %s", attempt + 1, str(e))
                last_error = e
                import asyncio
                await asyncio.sleep(2 ** attempt)

        raise RuntimeError(f"LLM call failed after {settings.max_retries} attempts: {last_error}")

    async def generate_text(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str | None = None,
        max_tokens: int = 8192,
    ) -> str:
        """Call LLM and return raw text response."""
        return await self._call_llm(system_prompt, user_prompt, model, max_tokens)

    async def _call_llm(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str | None = None,
        max_tokens: int = 8192,
    ) -> str:
        """Dispatch to the configured LLM provider."""
        if self._provider == "google":
            return await self._call_google(system_prompt, user_prompt, model, max_tokens)
        else:
            return await self._call_anthropic(system_prompt, user_prompt, model, max_tokens)

    async def _call_anthropic(
        self, system_prompt, user_prompt, model, max_tokens
    ) -> str:
        """Call Anthropic Claude API."""
        import anthropic

        if self._client is None:
            self._client = anthropic.AsyncAnthropic(api_key=self._api_key)

        target_model = model or self._model
        response = await self._client.messages.create(
            model=target_model,
            max_tokens=max_tokens,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )
        return response.content[0].text

    async def _call_google(
        self, system_prompt, user_prompt, model, max_tokens
    ) -> str:
        """Call Google Gemini API."""
        import google.generativeai as genai

        if self._client is None:
            genai.configure(api_key=self._api_key)
            self._client = True  # Mark as configured

        target_model = model or self._model
        gemini = genai.GenerativeModel(
            model_name=target_model,
            system_instruction=system_prompt,
            generation_config=genai.types.GenerationConfig(
                max_output_tokens=max_tokens,
                temperature=0.3,
            ),
        )

        # Gemini API is synchronous, run in thread pool
        import asyncio
        loop = asyncio.get_event_loop()
        response = await loop.run_in_executor(
            None,
            lambda: gemini.generate_content(user_prompt)
        )

        return response.text

    async def close(self):
        if self._provider == "anthropic" and self._client and hasattr(self._client, 'close'):
            await self._client.close()
        self._client = None


def _extract_json_block(text: str) -> str:
    """Extract JSON from markdown code blocks if present."""
    match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if match:
        return match.group(1).strip()

    text = text.strip()
    if text.startswith("{") or text.startswith("["):
        return text

    return text
