# XDR Architecture

## Overview

```
┌────────────┐   IQ stream    ┌──────────────────────┐  frames.bin   ┌──────────────────────────────┐
│  SDR HW    │ ─────────────► │  Python backend      │ ────────────► │  Excel (XDR.xlsm)           │
│ RTL/Airspy │   (SoapySDR)   │  FFT → FM → audio    │               │  OnTime loop + ring buffer  │
│ SDRPlay    │                │  writes frames.bin   │               │  paints cells, plays audio  │
└────────────┘                │  plays audio         │               │                             │
                              └─────────┬────────────┘               └──────────────▲──────────────┘
                                        │  cmd.bin (freq/sr/gain/refresh) ──────────┘
                                        │
                                    Samples: gen_iq.py (synthetic FM IQ, deterministic)
```

## Data files (project-local `xdr_data/`)

### frames.bin (Python → Excel)

Fixed-layout binary, one record per frame, overwritten in place:

| Offset | Size | Field |
|---|---|---|
| 0 | 4 | magic (`0x58445231` = "XDR1") |
| 4 | 4 | sequence number (uint32, increments per frame) |
| 8 | 8 | timestamp (unix seconds, float64) |
| 16 | 512 | 128 × float32 FFT magnitude bins (dB or linear) |
| 528 | 4 | peak bin index (uint32) |
| 532 | 4 | peak value (float32) |
| 536 | 4 | flags (bit 0 = source live/recorded, bit 1 = sample-rate valid) |
| 540 | 8 | sample rate (float64, Hz) |
| 548 | 4 | reserved (uint32) |
| **552** | | record size (total) |

### cmd.bin (Excel → Python)

Fixed 4-uint32 record (16 bytes): frequency Hz, sample rate, gain ×10, refresh ms.

## Excel engine

- `Application.OnTime` reentrant render loop (no Timer control)
- In-memory ring buffer of 64 frames
- One bulk `Interior.Color` array assignment per frame
- `Calculation=Manual` during stream; `ScreenUpdating` toggled only around paint
- Frequency/gain/refresh live in named ranges
- 64-bit-safe VBA (`LongPtr` discipline)

## Roadmap

| Phase | Deliverable | Reaction |
|---|---|---|
| 0 | Repo + workbook skeleton + helper boots | — |
| 1 | Fake SDR: random bins → CSV → cell waterfall | "Huh. This may actually work." |
| 2 | Live FFT from recorded IQ → binary protocol | "Oh no." |
| 3 | RTL-SDR / Airspy / SDRPlay via SoapySDR | "It's still working." |
| 4 | 250 ms real-time waterfall | "I should probably stop." |
| 5 | FM demodulation + audio | "WHY IS THIS ACTUALLY WORKING" |
| 6 | Feature parity with a bad SDR app | — |
| 7 | Cursed optimization (MMF!) | — |
| 8 | AM airband | "I've gone too far." |
| 9 | ADS-B | "This project has become sentient." |
| 10 | Release kit | — |