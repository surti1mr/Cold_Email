"""Runtime detection and writable paths for local vs Vercel serverless."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def is_serverless() -> bool:
    return bool(os.getenv("VERCEL") or os.getenv("AWS_LAMBDA_FUNCTION_NAME"))


def is_windows() -> bool:
    return sys.platform == "win32"


def project_dir() -> Path:
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    """Writable directory for uploads, sessions, and generated files."""
    if is_serverless():
        root = Path("/tmp/cold_email")
    else:
        root = project_dir()
    root.mkdir(parents=True, exist_ok=True)
    return root


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path
