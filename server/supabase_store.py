from __future__ import annotations

import json
import os
from urllib.parse import quote
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError


class SupabaseStore:
    """Small PostgREST-backed key/value store for app JSON documents."""

    def __init__(self) -> None:
        self.url = (os.environ.get("SUPABASE_URL") or "").rstrip("/")
        self.key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_KEY") or ""
        self.table = os.environ.get("SUPABASE_TABLE") or "app_data"
        self.timeout = float(os.environ.get("SUPABASE_TIMEOUT") or "10")
        self.last_error = ""

    @property
    def enabled(self) -> bool:
        return bool(self.url and self.key)

    def status(self) -> dict:
        return {
            "enabled": self.enabled,
            "url_set": bool(self.url),
            "key_set": bool(self.key),
            "table": self.table,
            "last_error": self.last_error,
        }

    def _endpoint(self, query: str = "") -> str:
        suffix = f"?{query}" if query else ""
        return f"{self.url}/rest/v1/{self.table}{suffix}"

    def _headers(self, extra: dict | None = None) -> dict:
        headers = {
            "apikey": self.key,
            "Authorization": f"Bearer {self.key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }
        if extra:
            headers.update(extra)
        return headers

    def _request(self, method: str, url: str, payload=None, headers: dict | None = None):
        data = None
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = Request(url, data=data, headers=self._headers(headers), method=method)
        try:
            with urlopen(req, timeout=self.timeout) as res:
                body = res.read().decode("utf-8")
                self.last_error = ""
                return json.loads(body) if body else None
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            self.last_error = f"HTTP {exc.code}: {detail or exc.reason}"
            raise
        except URLError as exc:
            self.last_error = str(exc.reason)
            raise

    def load(self, key: str):
        if not self.enabled:
            return None
        query = f"select=value&key=eq.{quote(str(key), safe='')}&limit=1"
        rows = self._request("GET", self._endpoint(query))
        if not rows:
            return None
        return rows[0].get("value")

    def save(self, key: str, value) -> None:
        if not self.enabled:
            return
        payload = [{"key": str(key), "value": value}]
        self._request(
            "POST",
            self._endpoint("on_conflict=key"),
            payload,
            {"Prefer": "resolution=merge-duplicates,return=minimal"},
        )

    def delete(self, key: str) -> None:
        if not self.enabled:
            return
        self._request("DELETE", self._endpoint(f"key=eq.{quote(str(key), safe='')}"), None, {"Prefer": "return=minimal"})

    def keys(self) -> list[str]:
        if not self.enabled:
            return []
        rows = self._request("GET", self._endpoint("select=key"))
        return [str(row.get("key")) for row in (rows or []) if row.get("key")]
