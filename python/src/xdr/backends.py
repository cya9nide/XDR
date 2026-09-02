"""SDR backend abstraction for XDR.

Priority: pyrtlsdr (pip-installable, RTL-SDR V3/V4, no extra SDK).
Fallback: SoapySDR (if installed + built with RTL/Airspy/SDRplay drivers) —
runtime-probed so absence never breaks the app.

Both backends expose the same DroneBackend interface:
    list_devices() -> list[str]          # human-readable name per device
    open_device(i) -> backend           # 0-indexed
    .set_freq(hz), .set_sample_rate(hz), .set_gain(db)
    .read_iq(n) -> np.ndarray[complex64]  # n IQ samples
    .close()
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np


class SDRNotFoundError(RuntimeError):
    pass


@dataclass
class SDRDevice:
    """One enumerable device, backend-agnostic."""
    index: int
    name: str
    backend: str  # 'pyrtlsdr' | 'soapy'


class _PyRtlsdrBackend:
    """pyrtlsdr (rtlsdr.RtlSdr) — pip-installable, RTL-SDR V3/V4."""

    name = "pyrtlsdr"

    def __init__(self) -> None:
        try:
            from rtlsdr import RtlSdr  # type: ignore
        except Exception as e:  # pragma: no cover
            raise SDRNotFoundError(f"pyrtlsdr not available: {e}") from e
        self._rtl = RtlSdr

    def list_devices(self) -> list[SDRDevice]:  # type: ignore[override]
        # pyrtlsdr exposes a single device at a time (first found).
        return [SDRDevice(0, "RTL-SDR (pyrtlsdr)", self.name)]

    def open_device(self, index: int = 0) -> "_PyRtlsdrBackend":  # type: ignore[override]
        try:
            self._dev = self._rtl()
        except Exception as e:  # pragma: no cover
            raise SDRNotFoundError(f"RTL-SDR open failed ({e}). Is the dongle plugged in & WinUSB driver set?") from e
        return self

    def set_freq(self, hz: int) -> None:
        self._dev.center_freq = hz

    def set_sample_rate(self, hz: int) -> None:
        self._dev.sample_rate = hz

    def set_gain(self, db: float) -> None:
        self._dev.gain = db

    def read_iq(self, n: int) -> np.ndarray:
        return self._dev.read_samples(n)  # complex64

    def close(self) -> None:
        if getattr(self, "_dev", None):
            self._dev.close()


class _SoapyBackend:
    """SoapySDR — broad driver support (RTL, Airspy, SDRplay) if installed."""

    name = "soapy"

    def __init__(self) -> None:
        try:
            import SoapySDR  # type: ignore
        except Exception as e:  # pragma: no cover
            raise SDRNotFoundError(f"SoapySDR not available: {e}") from e
        self._SoapySDR = SoapySDR
        self._dev = None

    def list_devices(self) -> list[SDRDevice]:
        try:
            devs = self._SoapySDR.Device.enumerate()
        except Exception as e:  # pragma: no cover
            raise SDRNotFoundError(f"SoapySDR enumerate failed: {e}") from e
        out = []
        for i, d in enumerate(devs):
            label = d.get("label") or d.get("driver") or f"device {i}"
            out.append(SDRDevice(i, f"{label} (SoapySDR)", self.name))
        return out

    def open_device(self, index: int = 0) -> "_SoapyBackend":  # type: ignore[override]
        try:
            devs = self._SoapySDR.Device.enumerate()
            self._dev = self._SoapySDR.Device(devs[index])
        except Exception as e:  # pragma: no cover
            raise SDRNotFoundError(f"SoapySDR open failed ({e})") from e
        return self

    def set_freq(self, hz: int) -> None:
        self._dev.setFrequency("RX", 0, hz)
        # best-effort correction for tuners that need it
        try:
            self._dev.setFrequency("RX", 0, hz, {"correction": 0})
        except Exception:
            pass

    def set_sample_rate(self, hz: int) -> None:
        self._dev.setSampleRate("RX", 0, hz)

    def set_gain(self, db: float) -> None:
        self._dev.setGain("RX", 0, db)

    def read_iq(self, n: int) -> np.ndarray:
        import SoapySDR  # type: ignore
        buf = np.empty(n, dtype=np.complex64)
        self._dev.readStream(self._stream, [buf], n)
        return buf

    def close(self) -> None:
        if self._dev is not None:
            try:
                self._dev.deactivateStream(self._stream)  # type: ignore[attr-defined]
            except Exception:
                pass
            try:
                self._dev.closeStream(self._stream)  # type: ignore[attr-defined]
            except Exception:
                pass
            self._dev = None


def probe() -> tuple[list[type], str]:
    """Return (backend_classes, preferred_name). Prefers pyrtlsdr, falls back to SoapySDR."""
    order: list[type] = []
    try:
        import rtlsdr  # noqa: F401
        order.append(_PyRtlsdrBackend)
    except Exception:
        pass
    try:
        import SoapySDR  # noqa: F401
        order.append(_SoapyBackend)
    except Exception:
        pass
    if not order:
        raise SDRNotFoundError("No SDR backend available: install pyrtlsdr or SoapySDR.")
    return order, order[0].name


def list_devices() -> tuple[list[SDRDevice], str]:
    """Enumerate devices across all available backends; prefer the present one.

    Returns (devices, backend_name_used). Raises SDRNotFoundError if none.
    """
    classes, _ = probe()
    for cls in classes:
        try:
            b = cls()
            devs = b.list_devices()
            if devs:
                return devs, cls.name
        except Exception:
            continue
    raise SDRNotFoundError("No SDR devices found on this system.")


def open_first() -> tuple[object, str]:
    """Open the first available device; returns (backend_instance, name)."""
    classes, _ = probe()
    for cls in classes:
        try:
            b = cls()
            b.open_device(0)
            return b, cls.name
        except Exception:
            continue
    raise SDRNotFoundError("No SDR device could be opened.")