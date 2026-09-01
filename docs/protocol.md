# XDR binary protocol

Fixed-layout, little-endian (Windows-native), no text parsing anywhere.

## frames.bin — Python → Excel

One record (552 bytes) per frame, overwritten in place at `xdr_data/frames.bin`:

| Offset | Size | Type | Field | Notes |
|---|---|---|---|---|
| 0    | 4  | u32  | magic     | `0x58445231` ("XDR1") |
| 4    | 4  | u32  | sequence  | increments per frame from Python |
| 8    | 8  | f64  | timestamp | unix seconds |
| 16   | 512| 128×f32 | bins   | FFT magnitude (dB), one per waterfall column |
| 528  | 4  | u32  | peak_idx  | bin with max magnitude |
| 532  | 4  | f32  | peak_val  | its magnitude (dB) |
| 536  | 4  | u32  | flags     | bit0: source live (1) / recorded (0); bit1: sample-rate valid |
| 540  | 8  | f64  | sample_rate | Hz |
| 548  | 4  | u32  | reserved  | 0 |

**Total record: 552 bytes.**

## cmd.bin — Excel → Python

Fixed 16-byte record of 4 × u32: `freq_hz, sample_rate, gain_x10, refresh_ms`.

- Python polls once per frame; zeros mean "no change".
- Excel writes on control-sheet `Worksheet_Change` (debounced).

## Versioning

Bump the magic + this doc together when layout changes (do not break consumers silently).