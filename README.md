# Excel Defined Radio (XDR)

> The world's first SDR application where the waterfall is rendered using spreadsheet cells.

**Motto:** *Because GNU Radio was too mainstream.*

**Official Support Statement:** No support will be provided. The existence of this software should not be interpreted as an endorsement of Excel as a radio platform.

**To future engineers:** We're sorry.

---

## What is this?

A functional software-defined radio where **Excel is the front end and the engine**. Python is a dumb peripheral that writes files; Excel reads them, paints a waterfall out of cell colors, and tells Python what to tune via a command file.

- Phase 0–2 (current): fake SDR → live FFT from recorded IQ → waterfall in cells
- Phase 3+: real hardware (RTL-SDR V3/V4, Airspy R2, SDRPlay via SoapySDR), FM demodulation, live audio
- Full roadmap: see `docs/architecture.md` and `H:\documents\XDR\plan-v1.1.md`

## Quick start (Phase 0)

1. Install Excel (any modern version, 32- or 64-bit).
2. `uv sync` in `/python` (or `pip install -r requirements.txt`).
3. `python main.py` → prints "Booting Excel-Defined Radio."
4. Open `excel/XDR.xlsm` → Enable Content.

## IPC (the cursed part)

Two fixed-layout binary files in a project-local `xdr_data/` dir:

- `frames.bin` — Python → Excel. 128-bin FFT frames, ~4–30 fps, overwritten in place.
- `cmd.bin` — Excel → Python. Tune / sample-rate / gain / refresh commands.

No ports. No services. No firewall prompts. A custom binary protocol whose only consumer is a spreadsheet.

## Project laws

1. Excel is the front end **and** the engine. If it doesn't happen in a workbook, it doesn't happen at all.
2. A stranger can run it. One clone, one sync, one `.xlsm`, zero admin.
3. Every phase is a shippable joke increment with a testable exit gate.
4. Keep the receipts — Diagnostics tab from day one.

## License

See `LICENSE`. (Spoiler: "No support will be provided.")