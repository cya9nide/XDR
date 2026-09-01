"""XDR signal processing — because GNU Radio was too mainstream.

Phase 2+: numpy FFT → magnitude bins. Phase 3+: real IQ from hardware.
"""

from __future__ import annotations

import numpy as np


def mag_bins(iq: np.ndarray, n_bins: int = 128) -> np.ndarray:
    """FFT an IQ array and return n_bins magnitude values (normalized dB).

    Guarded for zero/empty input. Returns 128 floats.
    """
    if iq is None or len(iq) == 0:
        return np.zeros(n_bins, dtype=np.float64)
    x = np.asarray(iq, dtype=np.complex128)
    n = len(x)
    # zero-pad / trim to a power-of-two window for a clean FFT
    nfft = max(2, 1 << int(np.ceil(np.log2(n))))
    if nfft != n:
        x = np.concatenate([x, np.zeros(nfft - n, dtype=np.complex128)])
    win = np.hanning(nfft)
    spectrum = np.fft.fftshift(np.abs(np.fft.fft(x * win))) + 1e-12
    # rebin nfft bins down to n_bins by averaging blocks
    block = nfft // n_bins
    if block >= 1:
        trimmed = spectrum[: block * n_bins].reshape(n_bins, block)
        mag = trimmed.mean(axis=1)
    else:  # n_bins > nfft: pad with zeros
        mag = np.zeros(n_bins)
        mag[:nfft] = spectrum
    db = 20.0 * np.log10(mag)
    # normalize to 0..1-ish scale, clip
    return np.clip((db - np.min(db)) / max(1e-9, np.max(db) - np.min(db)), 0.0, 1.0)