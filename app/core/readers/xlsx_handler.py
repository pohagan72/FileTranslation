"""XLSX read/translate implementation."""

from __future__ import annotations

import io
import logging
from typing import List, Tuple

import pandas as pd

from .base import DocumentHandler

logger = logging.getLogger(__name__)


class XlsxHandler(DocumentHandler):
    extension = ".xlsx"

    def collect_segments(self, stream: io.BytesIO) -> List[str]:
        segments, _ = self._scan(stream)
        return segments

    def apply_translations(self, stream: io.BytesIO, translations: List[str]) -> io.BytesIO:
        segments, df = self._scan(stream)
        if len(segments) != len(translations):
            raise RuntimeError(
                f"xlsx translation count mismatch: {len(segments)} segments "
                f"vs {len(translations)} translations"
            )

        # Walk the DataFrame in the same row-major order used by `_scan` and
        # substitute each translated cell.
        idx = 0
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                value = df.iat[r, c]
                if pd.notna(value):
                    text = str(value).strip()
                    if text:
                        df.iat[r, c] = translations[idx]
                        idx += 1

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output

    @staticmethod
    def _scan(stream: io.BytesIO) -> Tuple[List[str], pd.DataFrame]:
        stream.seek(0)
        df = pd.read_excel(stream)
        segments: List[str] = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                value = df.iat[r, c]
                if pd.notna(value):
                    text = str(value).strip()
                    if text:
                        segments.append(text)
        stream.seek(0)
        return segments, df
