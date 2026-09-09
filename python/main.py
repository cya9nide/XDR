"""XDR helper app — boots the cursed pipeline.

Modes:
  csv   — synthetic FFT from a CSV scarecrow file.
  bin   — read a pre-generated IQ file, FFT it, write binary frames.
  sdr   — live hardware: poll cmd.bin, stream FFT frames, demodulate audio.
"""

from __future__ import annotations

import argparse
import os
import queue
import time
from pathlib import Path

import numpy as np

from xdr.backends import SDRNotFoundError, list_devices, open_first
from xdr.dsp import mag_bins
from xdr.fm import fm_demod
from xdr.protocol import FLAG_LIVE, FLAG_SR_VALID, Frame, read_cmd, write_frame

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
    """Synthetic FFT: 128 values per line, overwritten in place."""
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
    """Read IQ file, FFT sliding windows, write binary frames in a loop."""
    out = DATA_DIR / "frames.bin"
    out.parent.mkdir(exist_ok=True)
    if not source.is_file():
        raise SystemExit(f"[xdr] live-fft requires an IQ file: {source}")
    iq = np.fromfile(source, dtype=np.complex64)
    if iq.size < 2:
        raise SystemExit(f"[xdr] IQ file too small or unreadable: {source}")
    print(f"[xdr] live-fft: {iq.size:,} complex64 samples from {source} ({iq.size / 2.4e6:.1f}s @ 2.4 MSPS)")
    from xdr.dsp import mag_bins

    n = 0
    step = 4096
    pos = 0
    t_last = 0.0
    try:
        while True:
            n += 1
            chunk = iq[pos : pos + step]
            if chunk.size < 2:
                pos = 0
                continue
            pos = (pos + step) % iq.size  # wrap: loop the file forever
            bins = mag_bins(chunk)
            # seq carries the real frame count; timestamp lets Excel measure cadence
            frm = Frame(n, bins, sample_rate=2_400_000, flags=FLAG_SR_VALID | FLAG_LIVE)
            write_frame(out, frm)
            if n % 15 == 0:
                now = time.perf_counter()
                fps = 15 / max(1e-9, now - t_last) if t_last else 0
                t_last = now
                print(f"[xdr] frame {n} peak={frm.peak_idx}@{frm.peak_val:.2f} ({fps:.1f} fps)")
            time.sleep(per_frame)
    except KeyboardInterrupt:
        print(f"\n[xdr] stopped after {n} frames")


def run_sdr(per_frame: float) -> None:
    """Live SDR — poll cmd.bin for tune, stream FFT frames, demodulate audio."""
    print("[xdr] sdr mode: probing backends...")
    try:
        devs, name = list_devices()
    except SDRNotFoundError as e:
        raise SystemExit(f"[xdr] {e}")
    print(f"[xdr] backend: {name}; devices: {[d.name for d in devs]}")
    try:
        dev, name = open_first()
    except SDRNotFoundError as e:
        raise SystemExit(f"[xdr] {e}")
    freq = 88_000_000
    sr = 2_400_000
    gain = 20.0
    try:
        dev.set_sample_rate(sr)
        dev.set_gain(gain)
        dev.set_freq(freq)
        print(f"[xdr] opened {name}: freq={freq/1e6:.2f} MHz sr={sr/1e6:.2f} MSPS gain={gain:.1f} dB")
    except Exception as e:
        print(f"[xdr] ! initial tune failed ({e}); continuing anyway")
    n = 0
    step = 4096          # FFT frame size
    audio_chunk = 120_000  # 50 ms IQ per audio emit
    out = DATA_DIR / "frames.bin"
    t_last = 0.0
    audio_stream = None
    audio_state: list | None = None   # de-emphasis filter state (across chunks)
    # phase-5 audio: BLOCKING write (measured best: 98.5%). The vectorized
    # stateful de-emphasis removes per-chunk compute jitter; blocksize 2048
    # gives the device headroom. (callback+queue was tried → worse: latency.)
    try:
        import sounddevice as sd
        audio_stream = sd.OutputStream(
            samplerate=48_000, channels=1, dtype="float32",
            blocksize=2048,
        )
        audio_stream.start()
        audio_state = [np.zeros(0)]   # vectorized de-emphasis tail state
        print("[xdr] audio out: sounddevice 48 kHz mono")
    except Exception as e:
        print(f"[xdr] ! audio output unavailable ({e}); waterfall only")
    carrier_hz = 0.0     # station offset from center; set when tuning
    try:
        while True:
            n += 1
            # poll cmd.bin (Excel controls) — new tune applies next frame
            cfg = read_cmd(DATA_DIR / "cmd.bin")
            if cfg is not None and cfg.freq_hz and cfg.freq_hz != freq:
                freq = cfg.freq_hz
                try:
                    dev.set_freq(freq)
                    print(f"[xdr] tune -> {freq/1e6:.2f} MHz")
                except Exception as e:
                    print(f"[xdr] ! tune failed ({e})")
            # audio: read a real chunk every frame — the blocking read/write
            # self-throttles at real-time rate (120k samples @2.4MSPS = 50 ms,
            # so ~20 blocks/sec = the audio device's own cadence). No sleep
            # gating → no buffer starvation → no choppiness.
            if audio_stream is not None:
                try:
                    big = dev.read_iq(audio_chunk)
                    audio = fm_demod(big, sample_rate=sr, carrier_hz=carrier_hz,
                                     state=audio_state)
                    if audio.size:
                        audio_stream.write(audio.astype(np.float32))
                except Exception as e:
                    print(f"[xdr] ! audio failed ({e})")
            # waterfall: FFT a slice of the same chunk (fast, no extra read)
            iq = big[:step] if audio_stream is not None else dev.read_iq(step)
            bins = mag_bins(iq)
            frm = Frame(n, bins, sample_rate=sr, flags=FLAG_SR_VALID | FLAG_LIVE)
            write_frame(out, frm)
            if n % 15 == 0:
                now = time.perf_counter()
                fps = 15 / max(1e-9, now - t_last) if t_last else 0
                t_last = now
                print(f"[xdr] frame {n} peak={frm.peak_idx}@{frm.peak_val:.2f} ({fps:.1f} fps)")
            # no time.sleep: the block read/write IS the pacing for audio;
            # waterfall rides along. For no-audio mode, gentle sleep keeps fps.
            if audio_stream is None:
                time.sleep(per_frame)
    except KeyboardInterrupt:
        print(f"\n[xdr] stopped after {n} frames")
    finally:
        if audio_stream is not None:
            try:
                audio_stream.stop()
                audio_stream.close()
            except Exception:
                pass
        dev.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--mode", choices=["fake", "live", "sdr"], default="fake")
    parser.add_argument("--source", type=Path, default=None, help="IQ file for --mode live")
    parser.add_argument("--wire", choices=["bin", "csv"], default="bin",
                        help="scarecrow (csv) or binary protocol (bin)")
    parser.add_argument("--fps", type=float, default=15.0, help="frames per second")
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
    elif args.mode == "sdr":
        run_sdr(per_frame)
    else:
        src = args.source or (Path(__file__).resolve().parent.parent / "samples" / "demo.iq")
        run_live_fft(src, per_frame)


if __name__ == "__main__":
    main()