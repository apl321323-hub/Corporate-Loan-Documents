from __future__ import annotations

import asyncio
import importlib.util
import os
import sys
from typing import Any

from firebase_functions import https_fn, options


BASE_DIR = os.path.dirname(__file__)
SERVER_DIR = os.path.join(BASE_DIR, "server")
if SERVER_DIR not in sys.path:
    sys.path.insert(0, SERVER_DIR)

os.environ.setdefault("APL_DATA_DIR", "/tmp/apl_data" if os.name != "nt" else os.path.join(BASE_DIR, ".runtime_data"))
os.environ.setdefault("SUPABASE_READ_FIRST", "1")

SERVER_MAIN_PATH = os.path.join(SERVER_DIR, "main.py")
server_spec = importlib.util.spec_from_file_location("apl_server_main", SERVER_MAIN_PATH)
if server_spec is None or server_spec.loader is None:
    raise RuntimeError("Could not load FastAPI server module")
server_module = importlib.util.module_from_spec(server_spec)
server_spec.loader.exec_module(server_module)
fastapi_app = server_module.app


def _header_bytes(headers) -> list[tuple[bytes, bytes]]:
    result: list[tuple[bytes, bytes]] = []
    for key, value in headers.items():
        if key.lower() == "host":
            continue
        result.append((str(key).lower().encode("latin-1"), str(value).encode("latin-1")))
    return result


def _response_headers(headers: list[tuple[bytes, bytes]]) -> dict[str, str]:
    result: dict[str, str] = {}
    for key, value in headers:
        name = key.decode("latin-1")
        text = value.decode("latin-1")
        if name.lower() in {"content-length"}:
            continue
        if name in result:
            result[name] = f"{result[name]}, {text}"
        else:
            result[name] = text
    return result


async def _call_asgi(req: https_fn.Request) -> tuple[int, dict[str, str], bytes]:
    body = req.get_data() or b""
    sent_body = False
    disconnect_event = asyncio.Event()
    status = 500
    headers: list[tuple[bytes, bytes]] = []
    chunks: list[bytes] = []

    path = req.path or "/"
    if not path.startswith("/"):
        path = f"/{path}"

    scope: dict[str, Any] = {
        "type": "http",
        "asgi": {"version": "3.0", "spec_version": "2.3"},
        "http_version": "1.1",
        "method": req.method,
        "scheme": "https",
        "path": path,
        "raw_path": path.encode("utf-8"),
        "query_string": (req.query_string or b""),
        "headers": _header_bytes(req.headers),
        "client": (req.remote_addr or "", 0),
        "server": (req.host or "", 443),
        "root_path": "",
    }

    async def receive() -> dict[str, Any]:
        nonlocal sent_body
        if not sent_body:
            sent_body = True
            return {"type": "http.request", "body": body, "more_body": False}
        # StreamingResponse listens for a disconnect in parallel with sending
        # the response body. Returning it immediately cancels the stream before
        # any XLSX bytes are emitted, so wait until Starlette cancels the task.
        await disconnect_event.wait()
        return {"type": "http.disconnect"}

    async def send(message: dict[str, Any]) -> None:
        nonlocal status, headers
        msg_type = message.get("type")
        if msg_type == "http.response.start":
            status = int(message.get("status") or 200)
            headers = list(message.get("headers") or [])
        elif msg_type == "http.response.body":
            chunks.append(message.get("body") or b"")

    await fastapi_app(scope, receive, send)
    return status, _response_headers(headers), b"".join(chunks)


@https_fn.on_request(
    region="asia-northeast3",
    timeout_sec=540,
    memory=options.MemoryOption.GB_2,
    secrets=[
        "SUPABASE_URL",
        "SUPABASE_SERVICE_ROLE_KEY",
        "SUPABASE_KEY",
        "SUPABASE_TABLE",
    ],
)
def api(req: https_fn.Request) -> https_fn.Response:
    if req.method == "OPTIONS":
        return https_fn.Response(
            "",
            status=204,
            headers={
                "Access-Control-Allow-Origin": "*",
                "Access-Control-Allow-Methods": "GET,POST,PUT,PATCH,DELETE,OPTIONS",
                "Access-Control-Allow-Headers": "Content-Type,Authorization",
            },
        )

    status, headers, body = asyncio.run(_call_asgi(req))
    if not any(key.lower() == "access-control-allow-origin" for key in headers):
        headers["Access-Control-Allow-Origin"] = "*"
    return https_fn.Response(body, status=status, headers=headers)
