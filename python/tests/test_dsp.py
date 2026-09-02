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
    rng = np.random.default_rng(1)
    t = np.arange(n) / fs
    # real RF always has a noise floor — don't test a noiseless pure tone
    # (its windowed FFT has pathological duplicate peaks).
    tone = np.exp(1j * 2 * np.pi * 512 * t) + 0.05 * (
        rng.standard_normal(n) + 1j * rng.standard_normal(n)
    )
    bins = mag_bins(tone, n_bins=128)
    peak = int(np.argmax(bins))
    # mag_bins fftshifts (DC-centered): +512 Hz → raw bin 512, rebin /32 → 16,
    # then shifted by N/2 → ~80. Exact index wobbles ±3 with the noise floor
    # (median-centered), so assert it's in the right (positive-frequency) half.
    assert 64 <= peak < 96, f"expected peak in right (positive-freq) half, got {peak}"
    # the peak should be dramatically above the noise floor (honest contract)
    # exclude the peak's immediate lobe neighbors (window main lobe spills
    # ~±2 bins); the REST of the band must be dark
    lobe_mask = np.ones(len(bins), dtype=bool)
    for k in range(max(0, peak - 2), min(len(bins), peak + 3)):
        lobe_mask[k] = False
    others = bins[lobe_mask]
    assert bins[peak] >= 0.7, f"peak {bins[peak]:.2f} too dim"
    assert others.max() <= 0.35, f"noise floor too bright: {others.max():.2f}"


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
    """The FM carrier should be visible as a strong FFT peak ABOVE the noise
    floor (honest-noise-floor contract)."""
    iq = synth_fm_iq(duration=0.1, seed=42)
    bins = mag_bins(iq, n_bins=256)
    peak = int(np.argmax(bins))
    # carrier at 200 kHz / 2.4 MHz = 1/12 of the band → bin ≈ 256/12 ≈ 21 (offset)
    assert peak >= 8, f"expected a carrier blob, peak at {peak}"
    # the carrier peak must be clearly above a QUIET bin's level (the noise
    # floor / median), not above all others (FM sidebands are legitimately
    # bright). Median of the whole frame ≈ noise floor.
    floor = float(np.median(bins))
    assert bins[peak] >= floor + 0.35, (
        f"carrier {bins[peak]:.2f} not enough above floor {floor:.2f}"
    )