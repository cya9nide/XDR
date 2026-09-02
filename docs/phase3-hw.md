# Phase 3 — Hardware bring-up checklist (RTL-SDR on Windows)

The moment you plug a dongle in, walk this list. It's short on purpose —
the software that reads it is already done.

## 0. Prereqs (done already)

- `python/` has `pyrtlsdr` (primary) + SoapySDR runtime-optional fallback.
- `main.py --mode sdr` streams FFT frames from the first device.
- Excel controls (B3 freq / B4 sample rate / B5 gain / B6 refresh) → `cmd.bin`
  → Python tunes live. Status panel rows 15-17 (Connected / Sample Rate / Signal).
- No dongle needed for the demo path (`--mode live` + `samples/demo.iq`).

## 1. Plug in the RTL-SDR

- USB port (USB2 is fine; RTL is a 2.0 device). Blinking LED = powered.
- Windows will usually pop "device not recognized" or install a generic driver.
  That's expected — it needs the WinUSB driver below.

## 2. Install the WinUSB driver (Zadig) — THE critical step

Realtek RTL2832U ships with a Microsoft "RTL2832U" / "Bulk-In, Interface" driver
that pyrtlsdr canNOT talk to. It needs **WinUSB** via Zadig.

1. Download Zadig: https://zadig.akeo.ie/
2. Run it (admin). Options → List All Devices.
3. Dropdown → **"Bulk-In, Interface (Interface 0)"** (or the RTL2832U entry;
   ALWAYS "Interface 0" for RTL-SDR, not "Interface 1").
4. Right-side target driver → **WinUSB** (not libusb, not usbser).
5. **Replace Driver** → wait for "Driver installed successfully".
6. Unplug/replug the dongle.

Verify: Device Manager → the dongle shows as **"WinUSB device"**, no yellow bang.
If it shows anything else, Zadig's step to WinUSB fixes most cases. Airspy needs
its own driver (see below), ignore that until the RTL works.

> 🚨 If you have TWO RTL-SDR dongles (e.g. V3 and V4), unplug the other while
> bringing up the first — pyrtlsdr binds the first device it finds.

## 3. Verify from Python

```bash
cd python
uv run python -c "from xdr import backends; print(backends.list_devices())"
```

Expect `[SDRDevice(index=0, name='RTL-SDR (pyrtlsdr)', backend='pyrtlsdr')]`.

> **Native driver:** `main.py` and `backends.py` automatically prepend
> `python/rtlsdr_libs/` (bundled x64 `rtlsdr.dll` + deps from the official
> rtl-sdr-blog Release) to PATH, so pyrtlsdr finds its native lib with **zero
> system installs**. If you still see "No SDR backend available", the DLLs
> aren't alongside — re-check `python/rtlsdr_libs/`.
> If it raises for another reason, the driver is wrong — redo step 2.

Then a 3-second live test:

```bash
uv run python -c "
from xdr import backends
d, name = backends.open_first()
d.set_sample_rate(2_400_000); d.set_gain(20); d.set_freq(98_000_000)
import numpy as np
iq = d.read_iq(4096)
print('IQ sample:', iq[0], 'std:', iq.std())
d.close()
"
```

std >> 0 and a sane complex sample = the dongle is feeding RF.

## 4. Drive it through Excel

1. `uv run python main.py --mode sdr --fps 15` (leave running)
2. Open `excel/XDR.xlsm`, Enable Content, `XDRMain.StartLoop`
3. B3 → 98.5 (MHz) → the Python console shows `tune -> 98.50 MHz`
4. Waterfall shows the FM band spectrum; B15 Connected → "YES"
   (set manually for now — Python status reporting comes with Phase 5)
5. Drag B3 around → the spectrum shifts live. That's the whole Phase 3 gate.

## 5. Optional: SoapySDR (Airspy / SDRplay)

Only needed for the non-RTL dongles. Install the SoapySDR SDK + airspy/sdrplay
modules from https://github.com/pothosware (or `vcpkg install soapysdr`),
then `pip install SoapySDR` — the module appears with the SDK. The `backends.py`
probe auto-detects it and lists all devices; no code change.

## Failure modes

| Symptom | Cause | Fix |
|---|---|---|
| `No RTL-SDR device found` | Wrong driver (still Bulk-In/Interface 1) | Zadig → WinUSB on Interface 0 |
| `Error -3` / `USB transfer failed` | Driver mismatch / unplugged | Re-Zadig; reseat cable |
| Chinese knock-off dongle | RTL-SDR V4 has a different tuner (R828D) | pyrtlsdr 0.5.0 handles V4; if not, use `rtl_test` to confirm |
| Two dongles, pyrtlsdr opens wrong one | — | unplug one; or add index selection later |
| Excel B3 edit → no tune | `cmd.bin` locked by antivirus | retry; check `err_cell` (B11) |

Reaction earned when this works: **"It's still working."** 🔥⚡🔥🤝