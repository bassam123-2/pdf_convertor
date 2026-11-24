# -*- coding: utf-8 -*-
"""Block every connection that is not this computer. The local page still works."""

from __future__ import annotations

import socket

_INSTALLED = False
_ORIGINAL = None

_LOOPBACK = {"127.0.0.1", "::1", "localhost"}


def _host_of(address: object) -> str:
    if isinstance(address, tuple) and address:
        return str(address[0] or "")
    return ""


def install() -> None:
    global _INSTALLED, _ORIGINAL
    if _INSTALLED:
        return
    _ORIGINAL = socket.socket.connect

    def _guarded(self, address):  # type: ignore[no-untyped-def]
        host = _host_of(address)
        if host not in _LOOPBACK:
            raise OSError("offline: this app does not use the network")
        return _ORIGINAL(self, address)

    socket.socket.connect = _guarded  # type: ignore[method-assign]
    _INSTALLED = True


def uninstall() -> None:
    global _INSTALLED, _ORIGINAL
    if _ORIGINAL is not None:
        socket.socket.connect = _ORIGINAL  # type: ignore[method-assign]
    _ORIGINAL = None
    _INSTALLED = False


def installed() -> bool:
    return _INSTALLED
