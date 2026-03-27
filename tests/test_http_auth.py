import time
from urllib.parse import parse_qs, urlparse

from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Route
from starlette.testclient import TestClient

import word_document_server.security.http_auth as http_auth
from word_document_server.security.http_auth import APIKeyMiddleware, build_download_signature


async def _ok_endpoint(request):
    return JSONResponse({"ok": True})


def _create_app() -> Starlette:
    app = Starlette(
        routes=[
            Route("/protected", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/health", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/healthz", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/files/{filename}", _ok_endpoint, methods=["GET", "OPTIONS"]),
        ]
    )
    app.add_middleware(
        APIKeyMiddleware,
        api_key="expected-secret",
        header_name="x-api-key",
        exempt_paths=["/health", "/healthz"],
    )
    return app


def test_missing_api_key_returns_401(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected")

    assert response.status_code == 401
    assert response.json() == {"error": "Missing API key header: x-api-key"}


def test_invalid_api_key_returns_403(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected", headers={"x-api-key": "wrong"})

    assert response.status_code == 403
    assert response.json() == {"error": "Invalid API key."}


def test_valid_api_key_passes(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected", headers={"x-api-key": "expected-secret"})

    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_health_paths_are_exempt(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response_health = client.get("/health")
    response_healthz = client.get("/healthz")

    assert response_health.status_code == 200
    assert response_healthz.status_code == 200


def test_options_passthrough(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.options("/protected")

    assert response.status_code == 200


def test_signed_file_url_valid_without_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    signature = build_download_signature("sample.docx", expires_at, "signing-secret")
    response = client.get(f"/files/sample.docx?exp={expires_at}&sig={signature}")

    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_signed_file_url_alias_params_are_accepted(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    signature = build_download_signature("sample.docx", expires_at, "signing-secret")
    response = client.get(f"/files/sample.docx?expires={expires_at}&signature={signature}")

    assert response.status_code == 200


def test_invalid_or_expired_signature_returns_403_even_with_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    response = client.get(
        f"/files/sample.docx?exp={expires_at}&sig=bad-signature",
        headers={"x-api-key": "expected-secret"},
    )

    assert response.status_code == 403
    assert response.json() == {"error": "Invalid or expired download signature."}


def test_file_route_without_signature_still_requires_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/files/sample.docx")

    assert response.status_code == 401
    assert response.json() == {"error": "Missing API key header: x-api-key"}


def test_build_download_url_signing_and_ttl(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)
    monkeypatch.delenv("DOC_DOWNLOAD_URL_TTL_SECONDS", raising=False)

    url = http_auth.build_download_url("https://example.com/files", "sample.docx")
    parsed = urlparse(url)
    params = parse_qs(parsed.query)

    assert parsed.path == "/files/sample.docx"
    assert "exp" in params
    assert params["exp"][0] == "1900"
    assert "sig" in params


def test_build_download_url_uses_minimum_ttl(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("WORD_MCP_API_KEY", raising=False)
    monkeypatch.setenv("DOC_DOWNLOAD_URL_TTL_SECONDS", "5")
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)

    url = http_auth.build_download_url("https://example.com/files", "sample.docx")
    params = parse_qs(urlparse(url).query)

    assert params["exp"][0] == "1030"


def test_build_download_url_falls_back_to_api_key_secret(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.setenv("WORD_MCP_API_KEY", "api-key-secret")
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)

    url = http_auth.build_download_url("https://example.com/files", "sample.docx")
    params = parse_qs(urlparse(url).query)
    expected_sig = build_download_signature("sample.docx", 1900, "api-key-secret")

    assert params["sig"][0] == expected_sig
