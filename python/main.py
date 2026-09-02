"""XDR helper app — boots the cursed pipeline.

Phase 1: fake SDR → CSV scarecrow → Excel paints cells.
Phase 2+: binary frames.bin wire. Phase 3+: SoapySDR hardware.
"""

from __future__ import annotations

import argparse
import os
import time
from pathlib import Path

import numpy as np

from xdr.protocol import FLAG_SR_VALID, Frame, write_frame

APP_TITLE = "Excel Defined Radio"


def default_data_dir() -> Path:
    env = os.environ.get("XDR_DATA_DIR")
    if env:
        return Path(env)
    return Path(__file__).resolve().parent.parent / "xdr_data"


DATA_DIR = default_data_dir()


def rng_bins(rng: np.random.Generator, n: int = 128, drift: float = 0.3) -> np.ndarray:
    """Random FFT-looking bins: noise floor + a drifting 'signal' blob."""
    base = rng.random(n) * 0.25
    center = int(n // 2 + 20 * np.sin(time.time() / 4))
    center = max(4, min(n - 5, center))
    width = 6
    for i in range(n):
        base[i] += 0.75 * np.exp(-((i - center) ** 2) / (2 * width**2))
    return base * (1 + drift * rng.random(n))


def run_fake(duration: float, per_frame: float) -> None:
    """Fake SDR, binary wire: write random bins to frames.bin."""
    out = DATA_DIR / "frames.bin"
    out.parent.mkdir(exist_ok=True)
    rng = np.random.default_rng(0xC0FFEE)
    n = 0
    t0 = time.perf_counter()
    deadline = t0 + duration if duration > 0 else None
    print(f"[xdr] fake bin mode: writing frames -> {out} ({per_frame:.2f}s/frame)")
    while True:
        n += 1
        bins = rng_bins(rng)
        write_frame(out, Frame(n, bins))
        if n % 10 == 0:
            print(f"[xdr] frame {n} seq={n} peak={bins.max():.2f}")
        if deadline and time.perf_counter() >= deadline:
            break
        time.sleep(per_frame)


def run_fake_csv(duration: float, per_frame: float) -> None:
    """Fake SDR, Phase 1 CSV scarecrow: 128 values per line, overwritten in place."""
    out = DATA_DIR / "frames.csv"
    out.parent.mkdir(exist_ok=True)
    rng = np.random.default_rng(0xC0FFEE)
    n = 0
    t0 = time.perf_counter()
    deadline = t0 + duration if duration > 0 else None
    print(f"[xdr] fake CSV mode: writing {out} ({per_frame:.2f}s/frame)")
    while True:
        n += 1
        bins = rng_bins(rng)
        line = ",".join(f"{b:.4f}" for b in bins)
        tmp = out.with_suffix(".tmp")
        tmp.write_text(line)
        # Windows: the target may be briefly locked (Excel FSO read, Defender scan).
        # Retry a short backoff; worst case a frame is skipped, never a crash.
        for _attempt in range(50):
            try:
                tmp.replace(out)
                break
            except PermissionError:
                time.sleep(0.01)
        else:
            out.write_text(line)  # final fallback: in-place write
        if n % 10 == 0:
            print(f"[xdr] csv frame {n} peak={bins.max():.2f}")
        if deadline and time.perf_counter() >= deadline:
            break
        time.sleep(per_frame)


def run_live_fft(source: Path, per_frame: float) -> None:
    """Phase 2: read IQ file, FFT, write frames. (Scaffold — filled next phase.)"""
    out = DATA_DIR / "frames.bin"
    out.parent.mkdir(exist_ok=True)
    if not source.is_file():
        raise SystemExit(f"[xdr] live-fft requires an IQ file: {source}")
    iq = np.fromfile(source, dtype=np.complex64)[: 1 << 20]
    if iq.size < 2:
        raise SystemExit(f"[xdr] IQ file too small or unreadable: {source}")
    print(f"[xdr] live-fft: {iq.size} samples from {source}")
    from xdr.dsp import mag_bins

    n = 0
    step = 4096
    pos = 0
    while True:
        n += 1
        chunk = iq[pos : pos + step]
        if chunk.size < 2:
            pos = 0
            continue
        pos = (pos + step) % iq.size
        bins = mag_bins(chunk)
        write_frame(out, Frame(n, bins, sample_rate=2_400_000, flags=FLAG_SR_VALID))
        if n % 10 == 0:
            print(f"[xdr] frame {n} peak={bins.max():.2f}")
        time.sleep(per_frame)


def main() -> None:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--mode", choices=["fake", "live"], default="fake")
    parser.add_argument("--source", type=Path, default=None, help="IQ file for --mode live")
    parser.add_argument("--wire", choices=["bin", "csv"], default="bin",
                        help="Phase 1 scarecrow (csv) or binary protocol (bin)")
    parser.add_argument("--fps", type=float, default=4.0, help="frames per second")
    parser.add_argument("--duration", type=float, default=0, help="seconds to run (0 = forever)")
    args = parser.parse_args()
    per_frame = 1.0 / args.fps if args.fps > 0 else 1.0
    print(f"{APP_TITLE} — booting.")
    print(f"  data dir: {DATA_DIR}")
    print(f"  mode: {args.mode}   wire: {args.wire}   fps: {args.fps}   duration: {args.duration}s")
    DATA_DIR.mkdir(exist_ok=True)
    if args.mode == "fake":
        if args.wire == "csv":
            run_fake_csv(args.duration, per_frame)
        else:
            run_fake(args.duration, per_frame)
    else:
        src = args.source or (Path(__file__).resolve().parent.parent / "samples" / "demo.iq")
        run_live_fft(src, per_frame)


if __name__ == "__main__":
    main()