"""FM demodulation — because GNU Radio was too mainstream.

Turn IQ samples into broadcast-audio samples with numpy.

Chain (all numpy, ~40 lines):
    1. freq_shift  — mix the carrier down to DC (center the station)
    2. decimate    — cut 2.4 MSPS to ~250 kSPS (audio band ~15 kHz)
    3. phase_diff  — FM demod: instantaneous frequency = dφ/dt
    4. de-emphasis — restore the broadcast pre-emphasis (75 µs US)
    5. downsample  — to 48 kHz for the audio device

Returns float64 mono audio at `audio_rate` Hz.
"""
from __future__ import annotations

import numpy as np


def fm_demod(
    iq: np.ndarray,
    sample_rate: float,
    carrier_hz: float,
    audio_rate: int = 48_000,
    decim: int = 10,
    tau: float = 75e-6,
    state: list | None = None,
) -> np.ndarray:
    """FM-demodulate `iq` (complex64/128) captured at `sample_rate`.

    `carrier_hz` is the station's offset from the tuned center frequency
    (i.e. the bin the station sits at, converted to Hz).

    `state`: optional 1-element list holding the de-emphasis filter's last
    output across calls — pass a persistent list (e.g. `[0.0]`) when feeding
    continuous chunks so the filter doesn't reset (and pop) at every boundary.
    """
    if iq is None or len(iq) == 0:
        return np.zeros(0, dtype=np.float64)
    x = np.asarray(iq, dtype=np.complex128)

    # 1. mix down: shift the station to DC
    n = len(x)
    t = np.arange(n) / sample_rate
    x = x * np.exp(-2j * np.pi * carrier_hz * t)

    # 2. decimate (low-pass + keep every `decim`th sample)
    #    FIR anti-alias: simple raised-cosine-ish window
    if decim > 1:
        # cheap but effective: boxcar-ish via reshape mean (decimate by averaging)
        keep = n - (n % decim)
        x = x[:keep].reshape(-1, decim).mean(axis=1)
        sample_rate /= decim

    # 3. FM demod: instantaneous frequency ∝ dφ/dt
    phase = np.angle(x)
    dphi = np.diff(phase)
    # unwrap jumps (|dφ| > π are phase wraps)
    dphi = np.mod(dphi + np.pi, 2 * np.pi) - np.pi
    # dphi already scaled: each sample is a phase step of 2π·f·dt → f = dphi/(2π dt)
    audio = dphi / (2 * np.pi) * sample_rate  # Hz deviation
    # normalize to a playable level (rough: broadcast FM ±75 kHz → ±1)
    audio = audio / 75_000.0

    # 4. de-emphasis (1st-order low-pass, 75 µs US) — vectorized, stateful.
    #    True IIR is a Python for-loop (slow, jittery). A first-order IIR's
    #    impulse response is h[k] = alpha·pole^k, so we approximate it as a
    #    truncated FIR via np.convolve (C-speed numpy). State = previous
    #    chunk's input tail, prepended as warmup so output stays seamless
    #    across chunk boundaries (no reset → no pop/chop).
    if tau > 0:
        alpha = float(1.0 - np.exp(-1.0 / (sample_rate * tau)))
        pole = 1.0 - alpha
        ntaps = 200  # 0.946^200 ≈ 2e-6 — enough for clean settling
        h = alpha * pole ** np.arange(ntaps)  # impulse response
        if state is not None and len(state[0]):
            tail = state[0]
            x_ext = np.concatenate([tail, audio])
            y_ext = np.convolve(x_ext, h, mode="full")[: len(x_ext)]
            y = y_ext[len(tail):]          # drop the warmup region
        else:
            y = np.convolve(audio, h, mode="full")[: len(audio)]
        if state is not None:
            state[0] = audio[-ntaps:]      # carry the input tail forward
        audio = y

    # 5. downsample to audio_rate (linear interp to arbitrary length)
    out_n = int(len(audio) * audio_rate / sample_rate)
    if out_n > 0:
        audio = np.interp(
            np.linspace(0, len(audio) - 1, out_n),
            np.arange(len(audio)),
            audio,
        )

    return audio


def freq_offset_to_carrier_hz(offset_bins: float, n_bins: int, sample_rate: float) -> float:
    """Convert an FFT bin offset from DC to Hz (for the station's carrier)."""
    return offset_bins * sample_rate / n_bins