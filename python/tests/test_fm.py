"""FM demod tests — prove the chain recovers a known tone from synthetic FM.

This is the "test before touching the dongle" gate: if we can FM a 440 Hz
tone, demod it, and get ~440 Hz back, the chain is right; only then do we
point it at real hardware.
"""
import numpy as np
import pytest

from xdr.fm import fm_demod, freq_offset_to_carrier_hz


def _synth_fm(tone_hz: float = 440.0, fs: float = 2_400_000.0,
              fc: float = 200_000.0, dev: float = 50_000.0,
              dur: float = 0.5, seed: int = 0) -> np.ndarray:
    """FM-modulate a tone onto a carrier + a bit of noise (realistic)."""
    rng = np.random.default_rng(seed)
    n = int(fs * dur)
    t = np.arange(n) / fs
    msg = 0.8 * np.sin(2 * np.pi * tone_hz * t)
    phase = 2 * np.pi * fc * t + dev * np.cumsum(msg) / fs
    iq = np.exp(1j * phase).astype(np.complex64)
    iq += 0.05 * (rng.standard_normal(n) + 1j * rng.standard_normal(n))
    return iq


def test_fm_demod_recovers_tone():
    """Demod a synthetic FM tone → recovered audio should be ~440 Hz."""
    iq = _synth_fm(tone_hz=440.0)
    audio = fm_demod(iq, sample_rate=2_400_000.0, carrier_hz=200_000.0)
    assert audio.size > 0
    # find the dominant frequency in the recovered audio
    audio_fs = 48_000.0
    n = len(audio)
    # hann + FFT of a slice, find peak
    win = np.hanning(n)
    spec = np.abs(np.fft.rfft(audio * win))
    freqs = np.fft.rfftfreq(n, d=1.0 / audio_fs)
    peak = freqs[int(np.argmax(spec))]
    assert 400 <= peak <= 480, f"expected ~440 Hz, got {peak:.1f} Hz"


def test_fm_demod_empty_input():
    assert fm_demod(np.array([], dtype=complex), 2.4e6, 0).size == 0


def test_freq_offset_to_carrier():
    # 32-bin offset at 2.4 MSPS / 128 bins = 600 kHz
    assert freq_offset_to_carrier_hz(32, 128, 2_400_000) == pytest.approx(600_000)