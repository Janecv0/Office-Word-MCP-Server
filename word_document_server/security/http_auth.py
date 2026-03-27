"""HTTP auth middleware and signed download URL helpers."""

from __future__ import annotations

import hashlib
import hmac
import os
import secrets
import time
from pathlib import Path
from typing import Any, List, Optional
from urllib.parse import quote, unquote

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse, Response

API_KEY_ENV_VAR = "WORD_MCP_API_KEY"
API_KEY_HEADER_ENV_VAR = "WORD_MCP_API_KEY_HEADER"
DOWNLOAD_SIGNING_SECRET_ENV_VAR = "DOC_DOWNLOAD_SIGNING_SECRET"
DOWNLOAD_URL_TTL_ENV_VAR = "DOC_DOWNLOAD_URL_TTL_SECONDS"


def _clean_value(raw_value: Optional[str]) -> str:
    value = (raw_value or "").strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
        value = value[1:-1].strip()
    return value


def get_api_key() -> Optional[str]:
    api_key = _clean_value(os.environ.get(API_KEY_ENV_VAR))
    if not api_key:
        return None
    return api_key


def get_api_key_header_name() -> str:
    header_name = _clean_value(os.environ.get(API_KEY_HEADER_ENV_VAR))
    if not header_name:
        return "x-api-key"
    return header_name.lower()


def get_download_signing_secret() -> Optional[str]:
    secret = _clean_value(os.environ.get(DOWNLOAD_SIGNING_SECRET_ENV_VAR))
    if secret:
        return secret
    return get_api_key()


def get_download_url_ttl_seconds() -> int:
    raw_value = _clean_value(os.environ.get(DOWNLOAD_URL_TTL_ENV_VAR)) or "900"
    try:
        ttl = int(raw_value)
    except (TypeError, ValueError):
        return 900
    return max(30, ttl)


def build_download_signature(filename: str, expires_at: int, secret: str) -> str:
    payload = f"{filename}:{expires_at}".encode("utf-8")
    return hmac.new(secret.encode("utf-8"), payload, hashlib.sha256).hexdigest()


def _extract_download_filename_from_path(path_value: str) -> Optional[str]:
    prefix = "/files/"
    if not path_value.startswith(prefix):
        return None

    raw_filename = path_value[len(prefix):]
    filename = _clean_value(unquote(raw_filename))
    if not filename:
        return None
    if Path(filename).name != filename:
        return None
    if not filename.lower().endswith(".docx"):
        return None
    return filename


def evaluate_signed_download_request(request: Request) -> str:
    """Validate signed /files/* URLs.

    Returns:
      - "valid" if signature is valid
      - "invalid" if signature params were attempted but invalid/expired
      - "not_attempted" if no signature params were provided
    """
    secret = get_download_signing_secret()
    if not secret:
        return "not_attempted"

    exp_raw = request.query_params.get("exp")
    if exp_raw is None:
        exp_raw = request.query_params.get("expires")

    signature = request.query_params.get("sig")
    if signature is None:
        signature = request.query_params.get("signature")

    if exp_raw is None and signature is None:
        return "not_attempted"

    filename = _extract_download_filename_from_path(request.url.path)
    if not filename:
        return "invalid"

    try:
        expires_at = int(exp_raw or "")
    except (TypeError, ValueError):
        return "invalid"

    if expires_at < int(time.time()):
        return "invalid"

    expected = build_download_signature(filename, expires_at, secret)
    if not signature or not secrets.compare_digest(signature, expected):
        return "invalid"

    return "valid"


def build_download_url(base_url: str, filename: str) -> str:
    """Build plain or signed download URL using existing base-url behavior."""
    normalized_base = (base_url or "").strip().rstrip("/")
    safe_name = quote(Path(filename).name)
    url = f"{normalized_base}/{safe_name}"

    secret = get_download_signing_secret()
    if not secret:
        return url

    expires_at = int(time.time()) + get_download_url_ttl_seconds()
    signature = build_download_signature(Path(filename).name, expires_at, secret)
    return f"{url}?exp={expires_at}&sig={signature}"


class APIKeyMiddleware(BaseHTTPMiddleware):
    """API-key protection middleware for streamable HTTP transport."""

    def __init__(
        self,
        app: Any,
        api_key: str,
        header_name: str = "x-api-key",
        exempt_paths: Optional[List[str]] = None,
    ):
        super().__init__(app)
        self.api_key = api_key
        self.header_name = (header_name or "x-api-key").strip().lower() or "x-api-key"
        self.exempt_paths = exempt_paths or []

    async def dispatch(self, request: Request, call_next: Any) -> Response:
        request_path = request.url.path

        if request.method.upper() == "OPTIONS":
            return await call_next(request)

        if request_path.startswith("/files/"):
            signed_status = evaluate_signed_download_request(request)
            if signed_status == "valid":
                return await call_next(request)
            if signed_status == "invalid":
                return JSONResponse(
                    {"error": "Invalid or expired download signature."},
                    status_code=403,
                )

        for exempt in self.exempt_paths:
            if request_path == exempt or request_path.startswith(f"{exempt.rstrip('/')}/"):
                return await call_next(request)

        provided_key = request.headers.get(self.header_name)
        if not provided_key:
            return JSONResponse(
                {"error": f"Missing API key header: {self.header_name}"},
                status_code=401,
            )

        if not secrets.compare_digest(provided_key, self.api_key):
            return JSONResponse({"error": "Invalid API key."}, status_code=403)

        return await call_next(request)
