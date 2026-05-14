from __future__ import annotations

import re


def safe_filename(text: str, max_len: int = 120) -> str:
    text = re.sub(r"[\\/:*?\"<>|]+", "_", str(text))
    text = re.sub(r"\s+", " ", text).strip()
    return text[:max_len].rstrip(" .") or "document"
