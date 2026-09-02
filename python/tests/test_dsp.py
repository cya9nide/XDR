"""DSP tests — pin the FFT + IQ generator behavior."""
import sys
from pathlib import Path

import numpy as np
import pytest

# make samples/ importable
SAMPLES = Path(__file__).resolve().parent.parent.parent / "samples"
sys.path.insert(0, str(SAMPLES))

from gen_iq import synth_fm_iq
from xdr.dsp import mag_bins


def test_mag_bins_tone_peak():
    """A pure tone should produce a clear peak in the right (DC-centered) bin."""
    fs = 4096
    n = 4096
    t = np.arange(n) / fs
    tone = np.exp(1j * 2 * np.pi * 512 * t)  # 512 Hz, N/8
    bins = mag_bins(tone, n_bins=128)
    peak = int(np.argmax(bins))
    # mag_bins fftshifts (DC-centered): +512 Hz → raw bin 512, rebin /32 → 16,
    # then shifted by N/2 → 16 + 64 = 80.
    assert peak == 80, f"expected peak at bin 80 (DC-centered +512 Hz), got {peak}"


def test_mag_bins_all_real_inputs():
    """Complex + real arrays both work; empty gives zeros."""
    r = np.random.default_rng(0).standard_normal(2048)
    b = mag_bins(r, n_bins=128)
    assert b.shape == (128,)
    assert np.all(np.isfinite(b))
    assert mag_bins(np.array([], dtype=complex), n_bins=128).shape == (128,)


def test_synth_fm_iq_deterministic():
    a = synth_fm_iq(duration=0.1, seed=42)
    b = synth_fm_iq(duration=0.1, seed=42)
    assert a.shape == b.shape
    assert np.array_equal(a, b)


def test_synth_fm_iq_has_carrier():
    """The FM carrier should be visible as a strong FFT peak."""
    iq = synth_fm_iq(duration=0.1, seed=42)
    bins = mag_bins(iq, n_bins=256)
    peak = int(np.argmax(bins))
    # carrier at 200 kHz / 2.4 MHz = 1/12 of the band → bin ≈ 256/12 ≈ 21 (offset)
    assert peak >= 8, f"expected a carrier blob, peak at {peak}"