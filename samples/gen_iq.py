"""Generate a deterministic synthetic FM-broadcast IQ file (Phase 2 demo source).

No hardware required: this makes an MVP demo always work, forever.

Output: interleaved complex64 IQ at FS sample rate; a strong FM-modulated
carrier offset from DC plus AWGN, so the FFT/waterfall shows a real,
moving, music-shaped signal.

Deterministic: same seed -> identical bytes (test pins this).
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np

FS = 2_400_000          # sample rate (Hz) — RTL-SDR-ish
SWEEP_AMP = 600_000     # carrier sweeps ±600 kHz (sinusoidal, 4s period)
SWEEP_PERIOD = 4.0      # full sweep cycle (seconds) — ~1 arc visible in 64-row history
DEVIATION = 120_000     # FM deviation (Hz per unit message) — visible wobble
NOISE = 0.02            # AWGN amplitude

# "music": (freq Hz, duration s) — a cheerful little arpeggio
NOTES = [
    (440.0, 0.15), (523.25, 0.10), (659.25, 0.15), (783.99, 0.10),
    (659.25, 0.15), (880.0, 0.20), (523.25, 0.15), (659.25, 0.10),
    (987.77, 0.20), (880.0, 0.15), (783.99, 0.10), (659.25, 0.30),
]


def synth_fm_iq(duration: float = 0.5, fs: int = FS, seed: int = 0xC0FFEE) -> np.ndarray:
    """Return complex64 IQ samples of a noise-floor + FM'd carrier."""
    rng = np.random.default_rng(seed)
    n = int(duration * fs)
    t = np.arange(n, dtype=np.float64) / fs

    # FM message: windowed notes stacked
    msg = np.zeros(n, dtype=np.float64)
    cursor = 0
    for freq, dur in NOTES:
        ln = int(dur * fs)
        if cursor + ln > n:
            break
        env = np.hanning(ln)
        msg[cursor : cursor + ln] += 0.8 * np.sin(2 * np.pi * freq * t[cursor : cursor + ln]) * env
        cursor += ln

    # FM modulate a SWEEPING carrier, so the waterfall shows visible motion:
    #   fc(t) = SWEEP_AMP * sin(2π t / SWEEP_PERIOD)  (glides across the band)
    #   phase_sweep = ∫ 2π fc dt = -SWEEP_AMP * SWEEP_PERIOD * cos(2π t / SWEEP_PERIOD)
    phase_sweep = -SWEEP_AMP * SWEEP_PERIOD * np.cos(2 * np.pi * t / SWEEP_PERIOD)
    phase_msg = DEVIATION * np.cumsum(msg) / fs
    phase = phase_sweep + phase_msg
    iq = np.exp(1j * phase).astype(np.complex64)

    # AWGN floor
    iq += (rng.standard_normal(n) + 1j * rng.standard_normal(n)).astype(np.complex64) * NOISE
    return iq


def main() -> None:
    parser = argparse.ArgumentParser(description="XDR synthetic FM IQ generator")
    parser.add_argument("--duration", type=float, default=0.5, help="seconds of IQ (default 0.5)")
    parser.add_argument("--seed", type=int, default=0xC0FFEE)
    parser.add_argument("--out", type=Path, default=Path(__file__).resolve().parent / "demo.iq")
    args = parser.parse_args()

    iq = synth_fm_iq(args.duration, seed=args.seed)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_bytes(iq.tobytes())
    print(f"wrote {args.out} ({iq.nbytes/1e6:.1f} MB, {iq.size:,} complex64 samples @ {FS/1e6:.1f} MSPS)")


if __name__ == "__main__":
    main()