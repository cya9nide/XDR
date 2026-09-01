"""Protocol round-trip tests — pin the binary contract."""
import struct
from pathlib import Path

import numpy as np

from xdr.protocol import (
    CMD_SIZE,
    FRAME_SIZE,
    MAGIC,
    Cmd,
    Frame,
    clear_cmd,
    read_cmd,
    read_frame,
    unpack_frame,
    write_cmd,
    write_frame,
)


def test_frame_roundtrip():
    bins = np.linspace(0.0, 1.0, 128)
    f = Frame(sequence=42, bins=bins, sample_rate=2_400_000, flags=0x3)
    data = f.pack()
    assert len(data) == FRAME_SIZE == 552
    f2 = unpack_frame(data)
    assert f2.magic == MAGIC
    assert f2.sequence == 42
    assert np.allclose(f2.bins, bins)
    assert f2.peak_idx == 127
    assert f2.sample_rate == 2_400_000
    assert f2.flags == 0x3


def test_frame_too_short():
    assert unpack_frame(b"") is None
    assert unpack_frame(struct.pack("<I", 0xDEAD)) is None


def test_frame_bad_magic():
    data = bytearray(b"\x00" * FRAME_SIZE)
    data[:4] = struct.pack("<I", 0x12345678)
    assert unpack_frame(bytes(data)) is None


def test_frame_file_roundtrip(tmp_path: Path):
    p = tmp_path / "frames.bin"
    bins = np.zeros(128, dtype=np.float64)
    bins[64] = 1.0
    write_frame(p, Frame(sequence=7, bins=bins))
    f = read_frame(p)
    assert f is not None
    assert f.sequence == 7
    assert f.peak_idx == 64


def test_read_frame_missing(tmp_path: Path):
    assert read_frame(tmp_path / "nope.bin") is None


def test_cmd_roundtrip(tmp_path: Path):
    p = tmp_path / "cmd.bin"
    c = Cmd(freq_hz=98_300_000, sample_rate=2_400_000, gain_x10=20, refresh_ms=250)
    write_cmd(p, c)
    c2 = read_cmd(p)
    assert (c2.freq_hz, c2.sample_rate, c2.gain_x10, c2.refresh_ms) == (
        98_300_000, 2_400_000, 20, 250,
    )
    assert len(c.pack()) == CMD_SIZE == 16


def test_cmd_bad_magic_zeros(tmp_path: Path):
    p = tmp_path / "cmd.bin"
    p.write_bytes(b"\x00" * 16)
    c = read_cmd(p)
    assert (c.freq_hz, c.sample_rate, c.gain_x10, c.refresh_ms) == (0, 0, 0, 0)


def test_clear_cmd(tmp_path: Path):
    p = tmp_path / "cmd.bin"
    write_cmd(p, Cmd(freq_hz=100_000_000))
    clear_cmd(p)
    c = read_cmd(p)
    assert (c.freq_hz, c.sample_rate, c.gain_x10, c.refresh_ms) == (0, 0, 0, 0)