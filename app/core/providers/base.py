"""Translation-provider abstraction.

Providers translate a single text segment into a target language; the
provider is expected to infer the source language itself. Adding a new
provider (OpenAI, DeepL, Azure) means implementing one class — the rest
of the application is provider-agnostic.
"""

from __future__ import annotations

from abc import ABC, abstractmethod


class TranslationError(RuntimeError):
    """Raised when a provider cannot translate (network, quota, blocked content)."""


class TranslationProvider(ABC):
    """Pure-logic translation interface — no Flask, no I/O frameworks."""

    @abstractmethod
    def translate(self, text: str, target_language: str) -> str:
        """Translate `text` and return the translated string.

        Implementations should raise `TranslationError` on failure rather than
        returning the original text — the caller decides how to recover.
        """
