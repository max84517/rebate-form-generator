"""Background job runner + event hub.

The pipeline modules take a ``log(msg, level)`` callback.  In the old
CustomTkinter UI that callback appended to a textbox; here it publishes to
every connected Server-Sent-Events listener instead.

Only one job runs at a time (mirroring the old ``_set_running`` behaviour).
"""
from __future__ import annotations

import json
import queue
import threading
import traceback
from typing import Any, Callable

_MAX_HISTORY = 2000


class Hub:
    """Fan-out event bus + single-slot job runner."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._listeners: list[queue.Queue] = []
        self._history: list[dict] = []
        self._seq = 0
        self._running: str | None = None
        self._ever_attached = False

    # ── subscriptions ────────────────────────────────────────────────
    def subscribe(self) -> queue.Queue:
        q: queue.Queue = queue.Queue()
        with self._lock:
            # replay history so a reloaded page still shows the current run
            for ev in self._history:
                q.put(ev)
            q.put({"type": "state", "running": self._running})
            self._listeners.append(q)
            self._ever_attached = True
        return q

    def unsubscribe(self, q: queue.Queue) -> None:
        with self._lock:
            if q in self._listeners:
                self._listeners.remove(q)

    @property
    def listener_count(self) -> int:
        """How many browser tabs are currently attached."""
        with self._lock:
            return len(self._listeners)

    @property
    def ever_attached(self) -> bool:
        """True once a browser has connected at least once.

        Recorded on subscribe rather than sampled, so a short-lived
        connection is never missed.
        """
        with self._lock:
            return self._ever_attached

    # ── publishing ───────────────────────────────────────────────────
    def publish(self, event: dict, remember: bool = False) -> None:
        with self._lock:
            self._seq += 1
            event = {**event, "seq": self._seq}
            if remember:
                self._history.append(event)
                if len(self._history) > _MAX_HISTORY:
                    del self._history[: len(self._history) - _MAX_HISTORY]
            listeners = list(self._listeners)
        for q in listeners:
            q.put(event)

    def log(self, msg: str, level: str = "INFO") -> None:
        self.publish({"type": "log", "level": level, "msg": str(msg)}, remember=True)

    def clear_history(self) -> None:
        with self._lock:
            self._history.clear()
        self.publish({"type": "clear"})

    # ── job control ──────────────────────────────────────────────────
    @property
    def running(self) -> str | None:
        with self._lock:
            return self._running

    def start(self, name: str, fn: Callable[[Callable[[str, str], None]], Any]) -> bool:
        """Run *fn* on a worker thread.  Returns False if a job is already running.

        *fn* receives the log callback and returns a JSON-serialisable result,
        which is delivered to the browser in the ``done`` event.
        """
        with self._lock:
            if self._running is not None:
                return False
            self._running = name

        self.publish({"type": "state", "running": name})

        def worker() -> None:
            try:
                result = fn(self.log)
                self.publish({"type": "done", "job": name, "result": result})
            except Exception as exc:  # noqa: BLE001 - surfaced to the UI
                self.log(f"Error: {exc}", "ERROR")
                self.log(traceback.format_exc(), "ERROR")
                self.publish({"type": "done", "job": name, "result": None,
                              "error": str(exc)})
            finally:
                with self._lock:
                    self._running = None
                self.publish({"type": "state", "running": None})

        threading.Thread(target=worker, daemon=True, name=f"job-{name}").start()
        return True


def sse_format(event: dict) -> bytes:
    return f"data: {json.dumps(event, ensure_ascii=False, default=str)}\n\n".encode("utf-8")
