#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import tempfile
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from template_export import export_ppt


DEFAULT_TEMPLATE = Path(
    "/Users/airtonalmeida/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized/AWR Template.potx"
)


class AwrHandler(SimpleHTTPRequestHandler):
    server_version = "AWRTemplateServer/1.0"

    def _send_json(self, code: int, payload: dict):
        raw = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

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
        return super().do_GET()

    def do_POST(self):
        if self.path == "/api/template":
            try:
                clen = int(self.headers.get("Content-Length", "0"))
                body = self.rfile.read(clen) if clen > 0 else b"{}"
                payload = json.loads(body.decode("utf-8"))
                if not isinstance(payload, dict):
                    raise ValueError("Invalid JSON payload.")
                name = str(payload.get("filename") or "uploaded_template.potx")
                b64 = payload.get("content_b64")
                if not isinstance(b64, str) or not b64:
                    raise ValueError("Missing content_b64.")
                lower = name.lower()
                if not (lower.endswith(".potx") or lower.endswith(".pptx")):
                    raise ValueError("Template must be .potx or .pptx")

                import base64
                import os

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
            clen = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(clen) if clen > 0 else b"{}"
            payload = json.loads(body.decode("utf-8"))
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
