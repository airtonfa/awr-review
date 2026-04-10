#!/usr/bin/env python3
from __future__ import annotations

import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

from http.server import ThreadingHTTPServer

from app_server import AwrHandler, DEFAULT_TEMPLATE


def app_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent


def pick_port(host: str = "127.0.0.1", start: int = 8080, attempts: int = 40) -> int:
    for p in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            if s.connect_ex((host, p)) != 0:
                return p
    raise RuntimeError("No free local port found for AWR Review app.")


def resolve_template() -> Path:
    env_path = os.environ.get("AWR_TEMPLATE_PATH")
    if env_path:
        p = Path(env_path).expanduser().resolve()
        if p.exists():
            return p
    if DEFAULT_TEMPLATE.exists():
        return DEFAULT_TEMPLATE.resolve()
    return Path("/tmp/awr_missing_template.potx")


def log_line(msg: str):
    try:
        log_dir = Path.home() / "Library" / "Logs" / "AWRReview"
        log_dir.mkdir(parents=True, exist_ok=True)
        with (log_dir / "launcher.log").open("a", encoding="utf-8") as f:
            f.write(msg.rstrip() + "\n")
    except Exception:
        pass


def open_browser_url(url: str):
    try:
        if webbrowser.open(url):
            log_line(f"Opened browser via webbrowser: {url}")
            return
    except Exception as e:
        log_line(f"webbrowser.open failed: {e}")
    try:
        subprocess.Popen(["open", url])
        log_line(f"Opened browser via macOS open: {url}")
    except Exception as e:
        log_line(f"Fallback open command failed: {e}")


def run():
    try:
        host = "127.0.0.1"
        port = pick_port(host=host, start=8080)
        root = app_root()
        template_path = resolve_template()
        log_line(f"Starting AWR Review launcher. root={root} template={template_path}")

        class Handler(AwrHandler):
            pass

        os.chdir(root)
        server = ThreadingHTTPServer((host, port), Handler)
        server.template_path = template_path  # type: ignore[attr-defined]

        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        log_line(f"HTTP server started at http://{host}:{port}")

        url = f"http://{host}:{port}"
        open_browser_url(url)

        # Try GUI control window; if tkinter isn't available, keep server running headless.
        try:
            from tkinter import BOTH, LEFT, RIGHT, Button, Label, StringVar, Tk  # local import on purpose

            window = Tk()
            window.title("AWR Review")
            window.geometry("560x120")
            window.resizable(False, False)
            status = StringVar(
                value=(
                    f"AWR Review is running on {url}\n"
                    f"Template: {template_path if template_path.exists() else 'not found (load from UI)'}"
                )
            )
            Label(window, textvariable=status, justify=LEFT, anchor="w").pack(fill=BOTH, padx=14, pady=12)

            def open_browser():
                open_browser_url(url)

            def quit_app():
                try:
                    server.shutdown()
                    server.server_close()
                finally:
                    window.destroy()

            Button(window, text="Open Browser", command=open_browser).pack(side=LEFT, padx=14, pady=8)
            Button(window, text="Quit", command=quit_app).pack(side=RIGHT, padx=14, pady=8)
            window.protocol("WM_DELETE_WINDOW", quit_app)
            log_line("Launching tkinter control window.")
            window.mainloop()
            return
        except Exception as e:
            log_line(f"tkinter unavailable; running headless. reason={e}")

        # Headless fallback for environments where tkinter is missing in runtime.
        while True:
            time.sleep(60)
    except Exception as e:
        log_line(f"Launcher failed: {e}")
        raise


if __name__ == "__main__":
    run()
