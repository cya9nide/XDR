# Excel Defined Radio (XDR)

> The world's first SDR application where the waterfall is rendered using spreadsheet cells.

**Motto:** *Because GNU Radio was too mainstream.*

**Official Support Statement:** No support will be provided. The existence of this software should not be interpreted as an endorsement of Excel as a radio platform.

**To future engineers:** We're sorry.

![Excel Defined Radio](assets/banner.png)

---

## What is this?

A functional software-defined radio where **Excel is the front end and the engine**. Python is a dumb peripheral that writes files; Excel reads them, paints a waterfall out of cell colors, and tells Python what to tune via a command file. No ports. No services. No firewall prompts. A custom binary protocol whose only consumer is a spreadsheet.

## Features

- **Live FM broadcast audio** — tune a real station from a spreadsheet cell (e.g. `B3 = 102.9`) and listen. Yes, really.
- **Waterfall in cells** — a 128×64 pixel display where every pixel is an Excel cell, colored by conditional formatting.
- **Real hardware** — RTL-SDR V3/V4 (Rafael Micro R828D) via pyrtlsdr. Airspy R2 / SDRPlay via SoapySDR (optional).
- **One-click start** — START ENGINE / STOP / STEP buttons on the SDR sheet. No Alt-F11, no macros manually run, no console.
- **Demo mode** — no dongle? `samples/gen_iq.py` makes a deterministic synthetic FM IQ file for a hardware-free waterfall.
- **Diagnostics tab** — receipts from day one.

## Quick start

1. Install Excel (any modern version).
2. `uv sync` in `/python` (or `pip install -r requirements.txt`).
3. Plug in an RTL-SDR (or run the demo with `samples/gen_iq.py`).
4. Open `excel/XDR.xlsm` → **Enable Content**.
5. Click **START ENGINE** → the waterfall comes alive.
6. Set `B3` to an FM frequency (try 102.9) → audio.

## IPC (the cursed part)

Two fixed-layout binary files in a project-local `xdr_data/` dir:

- `frames.bin` — Python → Excel. 128-bin FFT frames, ~4–30 fps, overwritten in place (552-byte XDR1 records).
- `cmd.bin` — Excel → Python. Tune / sample-rate / gain / refresh commands (16-byte XDRCMD).

## Project laws

1. Excel is the front end **and** the engine. If it doesn't happen in a workbook, it doesn't happen at all.
2. A stranger can run it. One clone, one sync, one `.xlsm`, zero admin.
3. Every phase is a shippable joke increment with a testable exit gate.
4. Keep the receipts — Diagnostics tab from day one.

## Gallery

Real hardware, real FM, real cells. (Full image pack: `H:\documents\XDR\assets-review\`)

- `assets/banner.png` — the banner.
- *(more images on the shared drive for review)*

## Roadmap

| Phase | State |
|---|---|
| 0–2 | fake SDR → live FFT → waterfall in cells |
| 3 | real RTL-SDR hardware + Excel tuning + status panel |
| 4 | polish (frequency scale, gradients) |
| 5 | **FM demod + live audio** *(the floor)* |
| 6 | presets, S-meter, scan, themes, favorites, recorder |
| 7 | cursed optimization (audio jitter, fps/cell/CPU metrics) |
| 8–10 | airband, ADS-B, release kit |

## License

See `LICENSE`. (Spoiler: "No support will be provided.")