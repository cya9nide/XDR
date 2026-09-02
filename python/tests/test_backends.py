"""Backend tests — probe/device listing behavior (no hardware required).

These must NOT touch a real dongle — they exercise the no-hardware path
(graceful SDRNotFoundError) and the module's structure.
"""
from pathlib import Path

import numpy as np
import pytest

from xdr import backends


def test_probe_no_hardware_raises_gracefully():
    """No SDR present → both backends unavailable → clean error message."""
    # Don't depend on machine state: probe() either raises SDRNotFoundError
    # (no backend), or returns the installed backend class list. We only
    # assert the error type/message shape when it raises.
    try:
        classes, name = backends.probe()
    except backends.SDRNotFoundError as e:
        assert "No SDR backend" in str(e)
        return
    assert classes, "probe returned no backend classes"


def test_list_devices_raises_or_returns():
    """list_devices either returns a list of SDRDevice or raises cleanly."""
    try:
        devs, name = backends.list_devices()
    except backends.SDRNotFoundError as e:
        # probe() raises "No SDR backend available" or list() raises "No SDR device..."
        assert "No SDR" in str(e)
        return
    assert devs, "list_devices returned empty device list"
    assert all(d.backend in ("pyrtlsdr", "soapy") for d in devs)


def test_open_first_raises_when_no_device():
    """open_first with no pluggable device → clean SDRNotFoundError."""
    try:
        dev, name = backends.open_first()
        dev.close()
    except backends.SDRNotFoundError as e:
        assert "could be opened" in str(e) or "No SDR" in str(e)