"""XDR binary protocol — the cursed wire format.

Fixed-layout, little-endian, no text parsing anywhere.
See docs/protocol.md for the full spec.

Layout is aligned so VBA can read it with a native Type (Get #f,, udt):
doubles land on 8-byte boundaries, singles/uint32 on 4-byte boundaries.

Offsets:
   0   4  magic      u32   ("XDR1" = 0x58445231)
   4   4  sequence   u32
   8   8  timestamp  f64   (unix seconds)
  16 512  bins       128 × f32 FFT magnitude (dB), one per waterfall column
 528   8  sample_rate f64  Hz
 536   4  peak_idx   u32   bin with max magnitude
 540   4  peak_val   f32   its magnitude
 544   4  flags      u32   bit0 live/recorded, bit1 sample-rate valid
 548   4  reserved   u32
 552      (record size)
"""

from __future__ import annotations

import struct
import time
from pathlib import Path

MAGIC = 0x58445231  # "XDR1"
FRAME_SIZE = 552     # bytes per frames.bin record
NBINS = 128           # waterfall columns
CMD_SIZE = 16         # bytes in cmd.bin

# flags bits (frames.bin)
FLAG_LIVE = 0x1        # source is live hardware (vs recorded)
FLAG_SR_VALID = 0x2    # sample_rate field is meaningful

_FRAME_STRUCT = struct.Struct("<I I d 128f d I f I I")
_CMD_STRUCT = struct.Struct("<IIII")


class Frame:
    __slots__ = (
        "magic", "sequence", "timestamp", "bins",
        "peak_idx", "peak_val", "flags", "sample_rate",
    )

    def __init__(
        self,
        sequence: int,
        bins,
        sample_rate: float = 0.0,
        flags: int = 0,
        timestamp: float | None = None,
    ):
        self.magic = MAGIC
        self.sequence = int(sequence)
        self.timestamp = time.time() if timestamp is None else timestamp
        b = list(bins)
        if len(b) > NBINS:
            b = b[:NBINS]
        elif len(b) < NBINS:
            b = b + [0.0] * (NBINS - len(b))
        self.bins = b
        self.peak_idx, self.peak_val = _peak(self.bins)
        self.flags = int(flags)
        self.sample_rate = float(sample_rate)

    def pack(self) -> bytes:
        return _FRAME_STRUCT.pack(
            self.magic,
            self.sequence,
            self.timestamp,
            *self.bins,
            self.sample_rate,
            self.peak_idx,
            self.peak_val,
            self.flags,
            0,  # reserved
        )


def _peak(bins) -> tuple[int, float]:
    if not bins:
        return 0, 0.0
    idx = max(range(len(bins)), key=lambda i: bins[i])
    return idx, float(bins[idx])


def unpack_frame(data: bytes):
    """Parse a 552-byte buffer (or a full file). Returns Frame or None."""
    if len(data) < FRAME_SIZE:
        return None
    (magic, seq, ts, *rest) = _FRAME_STRUCT.unpack(data[:FRAME_SIZE])
    if magic != MAGIC:
        return None
    bins = rest[:128]
    (sample_rate, peak_idx, peak_val, flags, _reserved) = rest[128:]
    f = Frame.__new__(Frame)
    f.magic = magic
    f.sequence = seq
    f.timestamp = ts
    f.bins = bins
    f.peak_idx = peak_idx
    f.peak_val = peak_val
    f.flags = flags
    f.sample_rate = sample_rate
    return f


def read_frame(path) -> Frame | None:
    """Read the latest frame from frames.bin. Returns None if missing/invalid."""
    p = Path(path)
    try:
        data = p.read_bytes()
    except (FileNotFoundError, OSError):
        return None
    return unpack_frame(data)


def write_frame(path, frame: Frame) -> None:
    """Atomically write one frame to frames.bin (temp + rename).

    Windows: the target may be briefly locked (a concurrent reader or AV scan).
    Retry with a short backoff rather than crashing.
    """
    import time as _time

    p = Path(path)
    tmp = p.with_suffix(".tmp")
    tmp.write_bytes(frame.pack())
    for _attempt in range(50):
        try:
            tmp.replace(p)
            return
        except PermissionError:
            _time.sleep(0.01)
    # last resort: direct overwrite (reader may catch a partial frame; guarded upstream)
    p.write_bytes(frame.pack())


# ── cmd.bin ───────────────────────────────────────────────────────────────

class Cmd:
    __slots__ = ("freq_hz", "sample_rate", "gain_x10", "refresh_ms")

    def __init__(self, freq_hz=0, sample_rate=0, gain_x10=0, refresh_ms=0):
        self.freq_hz = int(freq_hz)
        self.sample_rate = int(sample_rate)
        self.gain_x10 = int(gain_x10)
        self.refresh_ms = int(refresh_ms)
        # sanitize — nobody needs a 0-refresh spin (Excel side also guards)
        if self.refresh_ms < 30:
            self.refresh_ms = 250

    def pack(self) -> bytes:
        return _CMD_STRUCT.pack(
            self.freq_hz, self.sample_rate, self.gain_x10, self.refresh_ms
        )

    @classmethod
    def unpack(cls, data: bytes) -> "Cmd":
        if len(data) < CMD_SIZE:
            return cls()
        f, s, g, r = _CMD_STRUCT.unpack(data[:CMD_SIZE])
        return cls(f, s, g, r)


def read_cmd(path) -> Cmd:
    """Read cmd.bin (does not clear — the loop clears after consuming)."""
    try:
        return Cmd.unpack(Path(path).read_bytes())
    except (FileNotFoundError, OSError):
        return Cmd()


def write_cmd(path, cmd: Cmd) -> None:
    Path(path).write_bytes(cmd.pack())


def clear_cmd(path) -> None:
    """Zero the command file so stale commands aren't re-applied."""
    Path(path).write_bytes(Cmd().pack())