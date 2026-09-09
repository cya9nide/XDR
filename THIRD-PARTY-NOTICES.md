# Third-Party Notices

XDR redistributes and/or depends on the following third-party software. Their
respective licenses are included, and each item is credited below as required
by the applicable license.

## rtl-sdr-blog — RTL-SDR Blog drivers (`bin/rtl-sdr-blog/*`, including `rtlsdr.dll`, `rtl_fm.exe`, `rtl_sdr.exe`, the static libs, and the bundled runtime DLLs)

- **Project:** https://github.com/rtlsdrblog/rtl-sdr-blog
- **License:** GNU General Public License **v2.0** — see `licenses/GPL-2.0.txt`
- **Source:** the complete corresponding source for these binaries is available
  at https://github.com/rtlsdrblog/rtl-sdr-blog

These are modified Osmocom `/ librtlsdr` drivers distributed in binary form.
Under GPL-2.0 §3, the corresponding source is offered above.

## Osmocom / librtlsdr — `python/rtlsdr_libs/librtlsdr.dll`

- **Project:** https://github.com/osmocom/rtl-sdr (mirror: https://gitea.osmocom.org/sdr/rtl-sdr)
- **License:** GNU General Public License **v2.0** — see `licenses/GPL-2.0.txt`
- **Source:** https://github.com/osmocom/rtl-sdr

## pyrtlsdr — Python wrapper for librtlsdr (build dependency)

- **Project:** https://github.com/pyrtlsdr/pyrtlsdr
- **License:** GNU General Public License **v3.0-or-later** — see `licenses/GPL-3.0.txt`

## pyrtlsdrlib — pre-built librtlsdr installation helper (build dependency)

- **Project:** https://github.com/pyrtlsdr/pyrtlsdrlib
- **License:** MIT License — see `licenses/MIT.txt`

---

### Notes

- The bundled `bin/rtl-sdr-blog/` and `python/rtlsdr_libs/` binaries are
  third-party builds; their build tools may carry their own notices embedded
  in the binaries (e.g. MSVC/pthread runtime attribution). We redistribute them
  unmodified and decline any warranty.
- XDR itself is licensed under GPL-2.0-or-later (see `LICENSE`) so that it can
  be combined and redistributed with the GPL-2.0 third-party components above.
- If you re-distribute this repository, you must retain this notice, the
  `licenses/` texts, and the corresponding-source links.
