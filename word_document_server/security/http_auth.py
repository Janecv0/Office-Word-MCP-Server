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
from starlette.responses import HTMLResponse, JSONResponse, Response

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
      - "expired" if signature params were validly parsed but expired
      - "invalid" if signature params were attempted but invalid
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
        return "expired"

    expected = build_download_signature(filename, expires_at, secret)
    if not signature or not secrets.compare_digest(signature, expected):
        return "invalid"

    return "valid"


def _client_prefers_html(request: Request) -> bool:
    accept = (request.headers.get("accept") or "").lower()
    return "text/html" in accept


def _signed_link_error_response(request: Request, signed_status: str) -> Response:
    if _client_prefers_html(request):
        if signed_status == "expired":
            title = "Download Link Expired"
            message = "This link has expired. Please request a new download link."
            hint = "The previous URL is no longer valid for security reasons."
        else:
            title = "Invalid Download Link"
            message = "This download link is invalid. Please request a new link."
            hint = "Please check the full link or generate a fresh signed download URL."

        html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title}</title>
  <style>
    :root {{
      --brand-primary: #002554;
      --brand-primary-hover: #0e3a6f;
      --brand-aqua: #2cb1bc;
      --surface-app: #f7f8fa;
      --surface-raised: #ffffff;
      --surface-muted: #f1f3f6;
      --text-primary: #1a1f2b;
      --text-secondary: #4b5563;
      --border-default: #e2e8f0;
      --border-strong: #cbd5e1;
    }}
    @media (prefers-color-scheme: dark) {{
      :root {{
        --surface-app: #0b1526;
        --surface-raised: #0f1b33;
        --surface-muted: #152341;
        --text-primary: #e6e8eb;
        --text-secondary: #c7cdd8;
        --border-default: #1f2a37;
        --border-strong: #2a3646;
      }}
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: "Gothce", Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background:
        radial-gradient(1100px 420px at 90% -100px, rgba(44, 177, 188, 0.16), transparent 65%),
        radial-gradient(980px 360px at 10% -120px, rgba(0, 37, 84, 0.18), transparent 60%),
        var(--surface-app);
      color: var(--text-primary);
      display: grid;
      min-height: 100vh;
      place-items: center;
      padding: 22px;
    }}
    .card {{
      max-width: 560px;
      width: 100%;
      background: var(--surface-raised);
      border: 1px solid var(--border-default);
      border-radius: 14px;
      overflow: hidden;
      box-shadow: 0 18px 46px rgba(7, 14, 30, 0.2);
    }}
    .brand {{
      background: linear-gradient(135deg, var(--brand-primary) 0%, var(--brand-primary-hover) 100%);
      color: #ffffff;
      padding: 16px 22px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.2);
    }}
    .wordmark {{
      margin: 0;
      font-size: 0.92rem;
      line-height: 1.2;
      letter-spacing: 0.07em;
      text-transform: uppercase;
      font-weight: 600;
      opacity: 0.95;
    }}
    .product {{
      margin: 4px 0 0;
      font-size: 0.95rem;
      letter-spacing: 0.04em;
      opacity: 0.96;
    }}
    .content {{
      padding: 24px 22px 22px;
    }}
    .status {{
      display: inline-flex;
      align-items: center;
      gap: 8px;
      background: var(--surface-muted);
      border: 1px solid var(--border-strong);
      color: var(--text-secondary);
      border-radius: 999px;
      padding: 6px 10px;
      font-size: 0.78rem;
      letter-spacing: 0.03em;
      text-transform: uppercase;
      font-weight: 600;
    }}
    .dot {{
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--brand-aqua);
      box-shadow: 0 0 0 3px rgba(44, 177, 188, 0.22);
    }}
    h1 {{
      margin: 14px 0 10px;
      font-size: 1.4rem;
      line-height: 1.2;
      color: var(--text-primary);
    }}
    p {{
      margin: 0 0 10px;
      line-height: 1.55;
      color: var(--text-secondary);
      font-size: 0.98rem;
    }}
    .footer {{
      margin-top: 16px;
      padding-top: 14px;
      border-top: 1px solid var(--border-default);
      color: var(--text-secondary);
      font-size: 0.82rem;
      letter-spacing: 0.02em;
    }}
  </style>
</head>
<body>
  <main class="card">
    <header class="brand">
      <p class="wordmark">MATTONI 1873</p>
      <p class="product">DIGITAL | Mchat</p>
    </header>
    <section class="content">
      <span class="status"><span class="dot"></span> Secure Download</span>
      <h1>{title}</h1>
      <p>{message}</p>
      <p>{hint}</p>
      <p class="footer">Mattoni 1873 - M chat</p>
    </section>
  </main>
</body>
</html>"""
        return HTMLResponse(content=html, status_code=403)

    return JSONResponse(
        {"error": "Invalid or expired download signature."},
        status_code=403,
    )


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
            if signed_status in ("invalid", "expired"):
                return _signed_link_error_response(request, signed_status)

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
