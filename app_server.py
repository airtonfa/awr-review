#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import tempfile
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote
from urllib import error as urlerror
from urllib import request as urlrequest

from template_export import export_ppt


DEFAULT_TEMPLATE = Path(
    "/Users/airtonalmeida/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/AWR Template.potx"
)
CHATBOT_ENABLED = False


class AwrHandler(SimpleHTTPRequestHandler):
    server_version = "AWRTemplateServer/1.0"

    def _send_json(self, code: int, payload: dict):
        raw = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> dict:
        clen = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(clen) if clen > 0 else b"{}"
        payload = json.loads(body.decode("utf-8"))
        if not isinstance(payload, dict):
            raise ValueError("Invalid JSON payload.")
        return payload

    def _load_env_from_root(self) -> bool:
        env_path = Path(self.server.root_path) / ".env"  # type: ignore[attr-defined]
        if not env_path.exists():
            return False
        loaded = False
        for line in env_path.read_text(encoding="utf-8").splitlines():
            raw = line.strip()
            if not raw or raw.startswith("#") or "=" not in raw:
                continue
            key, value = raw.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
                loaded = True
        return loaded

    def _load_codex_auth_key(self) -> bool:
        auth_path = Path.home() / ".codex" / "auth.json"
        if not auth_path.exists():
            return False
        try:
            payload = json.loads(auth_path.read_text(encoding="utf-8"))
        except Exception:
            return False
        if not isinstance(payload, dict):
            return False
        key = str(payload.get("OPENAI_API_KEY") or "").strip()
        if not key:
            return False
        os.environ["OPENAI_API_KEY"] = key
        model = str(payload.get("OPENAI_MODEL") or "").strip()
        if model and "OPENAI_MODEL" not in os.environ:
            os.environ["OPENAI_MODEL"] = model
        return True

    @staticmethod
    def _is_openai_api_key(value: str) -> bool:
        return bool(value) and value.startswith("sk-")

    def _resolve_openai_credential(self) -> tuple[str, str]:
        key = os.environ.get("OPENAI_API_KEY", "").strip()
        source = "env"
        if key:
            return key, source
        loaded = self._load_env_from_root()
        key = os.environ.get("OPENAI_API_KEY", "").strip()
        source = ".env" if loaded and key else "missing"
        if key:
            return key, source
        loaded = self._load_codex_auth_key()
        key = os.environ.get("OPENAI_API_KEY", "").strip()
        source = "~/.codex/auth.json" if loaded and key else "missing"
        return key, source

    def _chat_status(self) -> dict:
        key, source = self._resolve_openai_credential()
        if not key:
            return {
                "connected": False,
                "source": source,
                "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
                "error": "OPENAI_API_KEY not found.",
            }
        if not self._is_openai_api_key(key):
            return {
                "connected": False,
                "source": source,
                "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
                "error": "Credential is not an OpenAI API key (expected prefix sk-).",
            }
        return {"connected": True, "source": source, "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini")}

    def _chat_completion(self, messages: list[dict], model: str) -> str:
        api_key, source = self._resolve_openai_credential()
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY not found. Add it to environment, .env, or ~/.codex/auth.json.")
        if not self._is_openai_api_key(api_key):
            raise RuntimeError(
                f"Credential from {source} is a session/auth token, not an OpenAI API key. "
                "Use a Platform API key starting with 'sk-'."
            )

        # Convert chat-style messages to Responses API input text.
        parts = []
        for msg in messages:
            role = str(msg.get("role") or "user").strip()
            content = str(msg.get("content") or "").strip()
            if not content:
                continue
            parts.append(f"{role.upper()}:\n{content}")
        prompt = "\n\n".join(parts).strip()
        if not prompt:
            raise ValueError("Empty chat input.")

        payload = {
            "model": model or os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
            "input": prompt,
        }
        req = urlrequest.Request(
            "https://api.openai.com/v1/responses",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            },
            method="POST",
        )
        try:
            with urlrequest.urlopen(req, timeout=90) as resp:
                raw = resp.read().decode("utf-8")
        except urlerror.HTTPError as err:
            detail = err.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"OpenAI request failed ({err.code}): {detail or err.reason}") from err
        except Exception as err:
            raise RuntimeError(f"OpenAI request failed: {err}") from err

        data = json.loads(raw)
        text = data.get("output_text")
        if isinstance(text, str) and text.strip():
            return text

        # Fallback parser for structured output items.
        out = []
        for item in data.get("output", []) or []:
            for c in item.get("content", []) or []:
                t = c.get("text")
                if isinstance(t, str) and t.strip():
                    out.append(t)
        final = "\n".join(out).strip()
        if not final:
            raise RuntimeError("OpenAI response had no text output.")
        return final

    def do_GET(self):
        if self.path == "/api/health":
            template_exists = self.server.template_path.exists()  # type: ignore[attr-defined]
            self._send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "template_path": str(self.server.template_path),  # type: ignore[attr-defined]
                    "template_exists": template_exists,
                },
            )
            return
        if self.path == "/api/chat-status":
            if not CHATBOT_ENABLED:
                self._send_json(HTTPStatus.SERVICE_UNAVAILABLE, {"ok": False, "error": "Chat assistant is disabled."})
                return
            status = self._chat_status()
            self._send_json(HTTPStatus.OK, {"ok": True, **status})
            return
        if self.path == "/api/reference-library":
            docs_root = Path(self.server.root_path) / "docs"  # type: ignore[attr-defined]
            if not docs_root.exists() or not docs_root.is_dir():
                self._send_json(HTTPStatus.OK, {"ok": True, "root": str(docs_root), "docs": []})
                return

            allowed_ext = {
                ".pdf",
                ".md",
                ".txt",
                ".json",
                ".csv",
                ".log",
                ".doc",
                ".docx",
                ".ppt",
                ".pptx",
                ".potx",
                ".xls",
                ".xlsx",
            }
            docs = []
            for p in sorted(docs_root.rglob("*")):
                if not p.is_file():
                    continue
                if p.suffix.lower() not in allowed_ext:
                    continue
                rel = p.relative_to(Path(self.server.root_path))  # type: ignore[attr-defined]
                rel_posix = rel.as_posix()
                docs.append(
                    {
                        "name": p.name,
                        "path": rel_posix,
                        "url": "/" + quote(rel_posix, safe="/-_.() "),
                        "size_bytes": p.stat().st_size,
                        "modified": datetime.fromtimestamp(p.stat().st_mtime).isoformat(timespec="seconds"),
                    }
                )

            self._send_json(HTTPStatus.OK, {"ok": True, "root": str(docs_root), "docs": docs})
            return
        return super().do_GET()

    def do_POST(self):
        if self.path == "/api/chat-assistant":
            if not CHATBOT_ENABLED:
                self._send_json(HTTPStatus.SERVICE_UNAVAILABLE, {"ok": False, "error": "Chat assistant is disabled."})
                return
            try:
                payload = self._read_json_body()
                messages = payload.get("messages")
                if not isinstance(messages, list):
                    raise ValueError("Missing messages list.")
                status = self._chat_status()
                if not status["connected"]:
                    raise RuntimeError("No OPENAI_API_KEY found. Load ~/.codex/auth.json, .env, or set environment variable.")
                model = str(payload.get("model") or status["model"])
                answer = self._chat_completion(messages, model)
                self._send_json(HTTPStatus.OK, {"ok": True, "answer": answer, "model": model})
            except Exception as err:
                self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(err)})
            return

        if self.path == "/api/template":
            try:
                payload = self._read_json_body()
                name = str(payload.get("filename") or "uploaded_template.potx")
                b64 = payload.get("content_b64")
                if not isinstance(b64, str) or not b64:
                    raise ValueError("Missing content_b64.")
                lower = name.lower()
                if not (lower.endswith(".potx") or lower.endswith(".pptx")):
                    raise ValueError("Template must be .potx or .pptx")

                import base64

                raw = base64.b64decode(b64)
                if len(raw) > 60 * 1024 * 1024:
                    raise ValueError("Template file too large (max 60MB).")
                tpl_dir = Path(tempfile.gettempdir()) / "awr_review_templates"
                tpl_dir.mkdir(parents=True, exist_ok=True)
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_name = "".join(c for c in Path(name).name if c.isalnum() or c in ("-", "_", "."))
                out = tpl_dir / f"{stamp}_{safe_name}"
                out.write_bytes(raw)
                self.server.template_path = out.resolve()  # type: ignore[attr-defined]
                self._send_json(
                    HTTPStatus.OK,
                    {
                        "ok": True,
                        "template_path": str(self.server.template_path),  # type: ignore[attr-defined]
                        "template_exists": True,
                    },
                )
            except Exception as err:
                self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(err)})
            return

        if self.path != "/api/export-template":
            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Not found"})
            return
        try:
            payload = self._read_json_body()
            report = payload.get("report") if isinstance(payload, dict) else None
            if not isinstance(report, dict):
                raise ValueError("Missing 'report' JSON object in request body.")

            template_path = self.server.template_path  # type: ignore[attr-defined]
            if not Path(template_path).exists():
                raise FileNotFoundError(f"Template not found: {template_path}")

            with tempfile.TemporaryDirectory(prefix="awr_tpl_export_") as td:
                td_path = Path(td)
                in_path = td_path / "report.json"
                out_path = td_path / f"AWR_Analysis_Template_{datetime.now().strftime('%Y-%m-%dT%H-%M-%S')}.pptx"
                in_path.write_text(json.dumps(report, ensure_ascii=False), encoding="utf-8")
                export_ppt(in_path, Path(template_path), out_path)
                raw = out_path.read_bytes()

            fname = f"AWR_Analysis_Template_{datetime.now().strftime('%Y-%m-%dT%H-%M-%S')}.pptx"
            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        except Exception as err:
            self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": str(err)})


def main():
    ap = argparse.ArgumentParser(description="Serve AWR web app + template-native PPT export API.")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--root", default=".")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    args = ap.parse_args()

    root = Path(args.root).resolve()
    template_path = Path(args.template).resolve()

    class Handler(AwrHandler):
        pass

    server = ThreadingHTTPServer((args.host, args.port), Handler)
    server.root_path = root  # type: ignore[attr-defined]
    server.template_path = template_path  # type: ignore[attr-defined]
    print(f"Serving {root} on http://{args.host}:{args.port}")
    print(f"Template: {template_path}")
    try:
        import os

        os.chdir(root)
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
