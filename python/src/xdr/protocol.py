"""XDR binary protocol — the cursed wire format.

Fixed-layout, little-endian, no text parsing anywhere.
See docs/protocol.md for the full spec.
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

_FRAME_STRUCT = struct.Struct("<I I d 128f I f I d I")
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
        self.bins = list(bins[:NBINS]) if len(bins) >= NBINS else list(bins) + [0.0] * (NBINS - len(bins))
        if not (0 <= len(self.bins) <= NBINS):
            raise ValueError(f"bins must be 0..{NBINS} long, got {len(self.bins)}")
        self.peak_idx, self.peak_val = _peak(self.bins)
        self.flags = int(flags)
        self.sample_rate = float(sample_rate)

    def pack(self) -> bytes:
        return _FRAME_STRUCT.pack(
            self.magic,
            self.sequence,
            self.timestamp,
            *self.bins,
            self.peak_idx,
            self.peak_val,
            self.flags,
            self.sample_rate,
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
    (peak_idx, peak_val, flags, sr, _reserved) = rest[128:]
    f = Frame.__new__(Frame)
    f.magic = magic
    f.sequence = seq
    f.timestamp = ts
    f.bins = bins
    f.peak_idx = peak_idx
    f.peak_val = peak_val
    f.flags = flags
    f.sample_rate = sr
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
    """Atomically write one frame to frames.bin (temp + rename)."""
    p = Path(path)
    tmp = p.with_suffix(".tmp")
    tmp.write_bytes(frame.pack())
    tmp.replace(p)


# ── cmd.bin —──────────────────────────────────────────────────────────────

class Cmd:
    __slots__ = ("freq_hz", "sample_rate", "gain_x10", "refresh_ms")

    def __init__(self, freq_hz=0, sample_rate=0, gain_x10=0, refresh_ms=0):
        self.freq_hz = int(freq_hz)
        self.sample_rate = int(sample_rate)
        self.gain_x10 = int(gain_x10)
        self.refresh_ms = int(refresh_ms)

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