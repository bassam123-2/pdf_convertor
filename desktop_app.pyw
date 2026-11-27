# -*- coding: utf-8 -*-
"""
PDF Converter for Windows: the same page as `streamlit run app.py`, in a
window of its own.

    PDF Converter.exe               open the app
    PDF Converter.exe --serve PORT  (internal) the page server
    PDF Converter.exe --selftest F  check the packaged app, write a report to F

Opening the app shows a small loading card, starts the page server on this
computer only (127.0.0.1), and opens the page in a Microsoft Edge app
window: no address bar, no tabs, the app's own icon and title. Closing the
window stops the server. A second start while the app is open brings up
another window on the same server instead of starting a new one.
"""

from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

FROZEN = getattr(sys, "frozen", False)
# The folder that holds app.py, core/, engine/ (the bundle, or this checkout).
BUNDLE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
if str(BUNDLE) not in sys.path:
    sys.path.insert(0, str(BUNDLE))

NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
# The page server stops this long after its last window closed.
IDLE_SECONDS = 25


# --------------------------------------------------------------------------
# Paths and settings


def _home() -> Path:
    from engine.paths import app_home

    home = app_home()
    if home is None:
        local = os.environ.get("LOCALAPPDATA", "").strip() or str(Path.home() / "AppData" / "Local")
        from engine.version import APP_NAME

        home = Path(local) / APP_NAME
    home.mkdir(parents=True, exist_ok=True)
    return home


def _log(message: str) -> None:
    try:
        folder = _home() / "logs"
        folder.mkdir(parents=True, exist_ok=True)
        with open(folder / "desktop.log", "a", encoding="utf-8") as handle:
            handle.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")
    except Exception:
        pass


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.bind(("127.0.0.1", 0))
        return int(probe.getsockname()[1])


def _answers(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/_stcore/health", timeout=1.5) as reply:
            return reply.status == 200
    except Exception:
        return False


def _server_command(port: int) -> list[str]:
    if FROZEN:
        return [sys.executable, "--serve", str(port)]
    return [sys.executable, str(Path(__file__).resolve()), "--serve", str(port)]


# --------------------------------------------------------------------------
# The page server (runs in its own process)


def serve(port: int) -> None:
    """Run the Streamlit page on 127.0.0.1:port until no window has been open for a while."""
    os.chdir(BUNDLE)  # .streamlit/config.toml (theme, upload size) is read from here
    if sys.stdout is None or sys.stderr is None:
        # A windowed program has no console; the server still prints.
        stream = open(_home() / "logs" / "server.log", "a", encoding="utf-8", buffering=1)
        sys.stdout = sys.stdout or stream
        sys.stderr = sys.stderr or stream
    os.environ.setdefault("PDF_TOOL_DESKTOP", "1")
    os.environ.setdefault("PDF_TOOL_HIDE_EXCEL", "1")
    from engine.host import prepare_server, server_environment

    os.environ.update(server_environment())
    prepare_server()
    threading.Thread(target=_stop_when_idle, daemon=True).start()
    from streamlit.web import bootstrap

    flags = {
        "server.port": port,
        "server.address": "127.0.0.1",
        "server.headless": True,
        "server.fileWatcherType": "none",
        "server.runOnSave": False,
        "browser.gatherUsageStats": False,
        "browser.serverAddress": "127.0.0.1",
        "global.developmentMode": False,
        "client.toolbarMode": "minimal",
        "client.showErrorDetails": True,
    }
    # The options must be loaded before the server starts; run() alone
    # leaves the port and the address at their defaults.
    bootstrap.load_config_options(flag_options=flags)
    bootstrap.run(str(BUNDLE / "app.py"), False, [], flags)


def _stop_when_idle() -> None:
    """End the server once its windows are closed (and none came back)."""
    seen_window = False
    idle_since = None
    while True:
        time.sleep(2)
        try:
            from streamlit.runtime import Runtime

            runtime = Runtime.instance()
            active = runtime._session_mgr.num_active_sessions()
        except Exception:
            continue
        if active:
            seen_window = True
            idle_since = None
            continue
        if not seen_window:
            # Nobody opened the page within two minutes: something failed.
            idle_since = idle_since or time.time()
            if time.time() - idle_since > 120:
                os._exit(0)
            continue
        idle_since = idle_since or time.time()
        if time.time() - idle_since > IDLE_SECONDS:
            os._exit(0)


# --------------------------------------------------------------------------
# The window


def _edge() -> str | None:
    candidates = []
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
        if base:
            candidates.append(Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe")
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"), os.environ.get("LOCALAPPDATA")):
        if base:
            candidates.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    for path in candidates:
        if path.is_file():
            return str(path)
    return None


def open_window(port: int, size: tuple[int, int, int, int] | None = None) -> subprocess.Popen | None:
    """An app window on the page: Edge (or Chrome) with no address bar; the default browser otherwise."""
    url = f"http://127.0.0.1:{port}/"
    browser = _edge()
    if browser is None:
        import webbrowser

        webbrowser.open(url)
        return None
    width, height, left, top = size or (1280, 860, 80, 40)
    profile = _home() / "window"
    command = [
        browser,
        f"--app={url}",
        f"--user-data-dir={profile}",
        f"--window-size={width},{height}",
        f"--window-position={left},{top}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-sync",
        "--disable-features=Translate,msEdgeSidebarV2,msHubApps,msEdgeShoppingUI",
        "--lang=ar",
    ]
    try:
        return subprocess.Popen(command, creationflags=NO_WINDOW)
    except Exception as exc:
        _log(f"window failed: {exc}")
        import webbrowser

        webbrowser.open(url)
        return None


# --------------------------------------------------------------------------
# The loading card


class Splash:
    """A small card while the server starts: icon, name, a moving bar."""

    def __init__(self) -> None:
        import tkinter as tk
        from tkinter import ttk

        from engine.version import PRODUCT_NAME, VERSION

        self.tk = tk
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.configure(bg="#142033")
        width, height = 440, 260
        screen_w, screen_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.screen = (screen_w, screen_h)
        self.root.geometry(f"{width}x{height}+{(screen_w - width) // 2}+{(screen_h - height) // 2}")
        try:
            self.root.iconbitmap(str(BUNDLE / "packaging" / "icon.ico"))
        except Exception:
            pass
        frame = tk.Frame(self.root, bg="#142033", highlightthickness=1, highlightbackground="#2B3B55")
        frame.pack(fill="both", expand=True)
        try:
            self.picture = tk.PhotoImage(file=str(BUNDLE / "packaging" / "icon.png")).subsample(3, 3)
            tk.Label(frame, image=self.picture, bg="#142033").pack(pady=(26, 8))
        except Exception:
            tk.Frame(frame, height=40, bg="#142033").pack()
        tk.Label(frame, text=PRODUCT_NAME, fg="#FFFFFF", bg="#142033", font=("Segoe UI Semibold", 18)).pack()
        self.status = tk.StringVar(value="جارٍ التشغيل…")
        tk.Label(frame, textvariable=self.status, fg="#B9C6D6", bg="#142033", font=("Segoe UI", 10)).pack(pady=(4, 12))
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Card.Horizontal.TProgressbar", troughcolor="#22324B", background="#D6A84C", bordercolor="#22324B", lightcolor="#D6A84C", darkcolor="#D6A84C")
        self.bar = ttk.Progressbar(frame, mode="indeterminate", length=260, style="Card.Horizontal.TProgressbar")
        self.bar.pack()
        self.bar.start(12)
        tk.Label(frame, text=f"v{VERSION}", fg="#5E6F86", bg="#142033", font=("Segoe UI", 8)).pack(side="bottom", pady=8)

    def window_size(self) -> tuple[int, int, int, int]:
        screen_w, screen_h = self.screen
        width = min(1360, max(1000, int(screen_w * 0.82)))
        height = min(920, max(700, int(screen_h * 0.86)))
        return width, height, max(0, (screen_w - width) // 2), max(0, (screen_h - height) // 2 - 10)

    def wait_for(self, ready, timeout: float = 120.0) -> bool:
        """Keep the card alive until ready() is true; False on timeout."""
        start = time.time()
        result = {"ok": False}

        def tick() -> None:
            if ready():
                result["ok"] = True
                self.root.quit()
                return
            if time.time() - start > timeout:
                self.root.quit()
                return
            self.root.after(300, tick)

        self.root.after(300, tick)
        self.root.mainloop()
        return result["ok"]

    def fail(self, text: str) -> None:
        from tkinter import messagebox

        self.bar.stop()
        self.root.withdraw()
        messagebox.showerror("PDF Converter", text)

    def close(self) -> None:
        try:
            self.root.destroy()
        except Exception:
            pass


# --------------------------------------------------------------------------
# Start


def launch() -> int:
    lock = _home() / "server.json"
    # The app is already open: one more window on the same server.
    try:
        known = json.loads(lock.read_text(encoding="utf-8"))
        if _answers(int(known.get("port", 0))):
            open_window(int(known["port"]))
            return 0
    except Exception:
        pass

    try:
        splash = Splash()
    except Exception as exc:
        _log(f"splash failed: {exc}")
        splash = None

    port = _free_port()
    from engine.host import server_environment

    env = server_environment()
    env["PDF_TOOL_DESKTOP"] = "1"
    env["PDF_TOOL_HIDE_EXCEL"] = "1"
    log_file = open(_home() / "logs" / "server.log", "a", encoding="utf-8")
    server = subprocess.Popen(
        _server_command(port), env=env, cwd=str(BUNDLE),
        stdout=log_file, stderr=subprocess.STDOUT, creationflags=NO_WINDOW,
    )
    lock.write_text(json.dumps({"port": port, "pid": server.pid}), encoding="utf-8")
    _log(f"server started on {port} (pid {server.pid})")

    ready = lambda: _answers(port) or server.poll() is not None  # noqa: E731
    if splash is not None:
        splash.wait_for(ready)
        size = splash.window_size()
    else:
        deadline = time.time() + 120
        while not ready() and time.time() < deadline:
            time.sleep(0.3)
        size = None

    if server.poll() is not None or not _answers(port):
        message = "تعذّر تشغيل البرنامج. التفاصيل في ملف السجل:\n" + str(_home() / "logs" / "server.log")
        _log("server did not start")
        if splash is not None:
            splash.fail(message)
            splash.close()
        try:
            server.kill()
        except Exception:
            pass
        return 1

    open_window(port, size)
    if splash is not None:
        splash.close()
    # The server ends itself once its windows are closed; wait for it.
    server.wait()
    try:
        lock.unlink()
    except Exception:
        pass
    _log("server stopped")
    return 0


def main() -> int:
    args = sys.argv[1:]
    if args[:1] == ["--serve"] and len(args) >= 2:
        serve(int(args[1]))
        return 0
    if args[:1] == ["--selftest"]:
        from engine.selftest import run

        return run(args[1] if len(args) > 1 else "")
    return launch()


if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()
    sys.exit(main())
