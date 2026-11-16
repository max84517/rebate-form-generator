"""Entry point: start the local web UI and open it in Chrome.

The server binds to 127.0.0.1 only.  Chrome is launched in app mode with a
dedicated profile directory so that (a) it looks like a standalone window and
(b) closing that window reliably terminates the launcher process, which in
turn shuts the server down.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import webbrowser
from pathlib import Path

DEFAULT_PORT = 8765
APP_NAME = "RebateFormGenerator"


def find_chrome() -> str | None:
    candidates = [
        os.environ.get("CHROME_PATH"),
        rf"{os.environ.get('ProgramFiles', '')}\Google\Chrome\Application\chrome.exe",
        rf"{os.environ.get('ProgramFiles(x86)', '')}\Google\Chrome\Application\chrome.exe",
        rf"{os.environ.get('LocalAppData', '')}\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if path and Path(path).is_file():
            return path
    return None


def chrome_profile_dir() -> Path:
    base = os.environ.get("LocalAppData") or str(Path.home())
    path = Path(base) / APP_NAME / "chrome-profile"
    path.mkdir(parents=True, exist_ok=True)
    return path


def launch_chrome(url: str) -> subprocess.Popen | None:
    chrome = find_chrome()
    if not chrome:
        return None
    args = [
        chrome,
        f"--app={url}",
        f"--user-data-dir={chrome_profile_dir()}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-background-mode",
        "--window-size=1100,820",
    ]
    try:
        return subprocess.Popen(args)
    except OSError:
        return None


def main() -> None:
    parser = argparse.ArgumentParser(prog="rebate-form-generator")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT,
                        help=f"preferred port (default {DEFAULT_PORT})")
    parser.add_argument("--browser", choices=("chrome", "default", "none"),
                        default="chrome", help="how to open the UI")
    args = parser.parse_args()

    from rebate_form_generator.web.server import (
        find_running_instance,
        pick_port,
        serve,
        start_idle_watchdog,
    )

    # If a previous window is still open, reuse it instead of starting a
    # second server on another port - two instances mean two different URLs
    # and one of them showing stale data.
    existing = find_running_instance(args.port)
    if existing is not None:
        url = f"http://127.0.0.1:{existing}/"
        print(f"  {APP_NAME} is already running at {url} - opening that window.")
        if args.browser == "chrome" and launch_chrome(url) is None:
            webbrowser.open(url)
        elif args.browser == "default":
            webbrowser.open(url)
        sys.exit(0)

    port = pick_port(args.port)
    httpd, app, _thread = serve(port)
    url = f"http://127.0.0.1:{port}/"

    print("=" * 56)
    print(f"  {APP_NAME} - {url}")
    print("=" * 56)

    proc: subprocess.Popen | None = None
    if args.browser == "chrome":
        proc = launch_chrome(url)
        if proc is None:
            print("  Chrome not found - falling back to the default browser.")
            webbrowser.open(url)
    elif args.browser == "default":
        webbrowser.open(url)

    if args.browser != "none":
        print("  Close the app window to stop the server.")
        start_idle_watchdog(app)
    else:
        print("  Press Ctrl+C to stop the server.")

    try:
        app.shutdown_event.wait()
    except KeyboardInterrupt:
        pass

    print("Shutting down...")
    if proc is not None and proc.poll() is None:
        proc.terminate()
    try:
        httpd.shutdown()
        httpd.server_close()
    except Exception:
        pass
    # Hard exit: SSE handler threads can still be parked in a blocking write,
    # and a server that refuses to die is what caused the stale-port problem.
    sys.stdout.flush()
    os._exit(0)


if __name__ == "__main__":
    main()
