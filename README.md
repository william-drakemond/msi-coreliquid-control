# MSI MPG Coreliquid Control Panel

A Windows GUI for monitoring and controlling the **MSI MPG Coreliquid K240 / K360** AIO liquid cooler using the [liquidctl](https://github.com/liquidctl/liquidctl) library.

![Python](https://img.shields.io/badge/python-3.11%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- **Live monitoring** — RPM and duty% for radiator fans, water block fan, and pump, auto-refreshed every 2 seconds
- **Fan mode presets** — Silent / Balance / Game / Smart, each applying preset speeds to all channels at once
- **Per-channel speed sliders** — Manually set Rad Fans, Water Block, and Pump independently
- **Fan curve editor** — Visual drag-and-drop temperature→duty curve per channel
- **CPU temp feed** — Manually set a CPU temperature to feed to the device so it can follow the fan curve
- **Settings persistence** — Slider positions, last mode, and CPU temp are saved to `settings.json` and restored on next launch
- **System tray** — App minimises to the taskbar notification area (bottom-right); right-click → Exit to quit

---

## Screenshots

> Monitor tab with live fan speeds and per-channel sliders  
> Fan Curve editor with drag-and-drop control points

---

## Requirements

- Windows 10/11
- Python 3.11+
- USB device using **HidUsb** driver (NOT WinUSB — see [Driver Setup](#driver-setup))

```
pip install liquidctl pillow pystray pyinstaller
```

---

## Installation & Usage

### Run from source

```bash
git clone https://github.com/YOUR_USERNAME/msi-coreliquid-control
cd msi-coreliquid-control
pip install liquidctl pillow pystray
python gui.py
```

### Build standalone .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "MSI Coreliquid Control" --icon icon.ico \
  --hidden-import "pystray._win32" --hidden-import "liquidctl.driver.msi" --hidden-import "hid" \
  gui.py
```

The exe will be in `dist/MSI Coreliquid Control.exe`. No Python installation required to run it.

### Auto-start with Windows

1. Press `Win+R`, type `shell:startup`, press Enter
2. Copy a shortcut to `MSI Coreliquid Control.exe` into that folder

---

## Driver Setup

> **Important:** The device must use the Windows **HidUsb** driver, not WinUSB.

The MSI MPG Coreliquid is a USB HID device. liquidctl's `MpgCooler` driver uses the HIDAPI backend and requires the device to be on the HidUsb driver stack.

**Do NOT use Zadig to install WinUSB** for this device — doing so moves the device to the `PyUsbBus` backend which the `MpgCooler` driver does not support, and liquidctl will no longer detect it.

If you accidentally installed WinUSB via Zadig:

1. Open **Device Manager** → find the device under `Universal Serial Bus devices`
2. Right-click → **Uninstall device** → check "Delete the driver software" → OK
3. Reboot — Windows will re-install HidUsb automatically
4. Verify with `liquidctl list` — the device should appear

---

## Known Limitations

| Limitation | Detail |
|---|---|
| **Windows only** | Uses HIDAPI via Windows HID stack. Not tested on Linux/macOS. |
| **K240 misidentified as K360** | liquidctl reports this as `MPG Coreliquid K360`. The K240 has 2 radiator fans; Fan 3 always reads 0 RPM. This is harmless. |
| **No automatic CPU temp** | `psutil` cannot read CPU temperatures on Windows. You must manually set the CPU temp slider in the Fan Curve tab and click "Start Feeding" for the device to follow a curve. |
| **Fan curves need continuous temp feed** | The device requires a temperature to be sent every few seconds. If feeding stops, fan curve behaviour is undefined. |
| **Pump minimum 30%** | Setting pump below 30% is blocked in the UI. Running the pump at 0% risks overheating. |
| **OLED not supported** | The OLED display on K360 models is not controlled by this app. |
| **Single device** | The app connects to the first detected Coreliquid device. Multiple devices are not supported. |

---

## Fan Channels

| UI Label | liquidctl channel | Notes |
|---|---|---|
| Rad Fans | `fans` | Controls Fan 1 + Fan 2 together. Fan 3 always 0 on K240. |
| Water Block | `waterblock-fan` | Small fan on the CPU pump head. |
| Pump | `pump` | Pump motor speed. Minimum 30% recommended. |

---

## Mode Presets

| Mode | Rad Fans | Water Block | Pump |
|---|---|---|---|
| Silent | 60% | 60% | 60% |
| Balance | 70% | 70% | 70% |
| Game | 90% | 90% | 100% |
| Smart | 80% | 80% | 75% |

---

## Patches to liquidctl's `msi.py`

Two patches were applied to `liquidctl/driver/msi.py` to make this work on Windows. If you update liquidctl these patches need to be re-applied.

### Patch 1 — Fix HID device detection on Windows

**File:** `site-packages/liquidctl/driver/msi.py` — `MpgCooler.probe()`

**Problem:** On Windows, the device exposes multiple HID interfaces. The original probe filter was too aggressive — it excluded all handles where `usage_page == 0x0001`, which on Windows is the only valid handle, so the device was never detected and `liquidctl list` returned nothing.

**Original code:**
```python
if handle.hidinfo["usage_page"] == EXTRA_USAGE_PAGE:
    return
yield from super().probe(handle, **kwargs)
```

**Patched code:**
```python
if handle.hidinfo["usage_page"] == EXTRA_USAGE_PAGE and handle.hidinfo.get("usage", 0) != 0:
    return
yield from super().probe(handle, **kwargs)
```

**Why:** The intent of the filter is to skip the secondary OLED USB interface (which has `usage_page=0x0001` AND a non-zero `usage`). On Windows the main control interface also has `usage_page=0x0001` but `usage=0`, so it must be allowed through.

---

### Patch 2 — Skip OLED firmware query on K240

**File:** `site-packages/liquidctl/driver/msi.py` — `MpgCooler.connect()`

**Problem:** During `connect()`, the driver calls `get_oled_firmware_version()` which sends command `0xF1` and waits for a response. The K240 has no OLED screen, so it never responds, causing a `liquidctl.error.Timeout: operation timed out` and preventing the app from connecting at all.

**Original code:**
```python
self._oled_firmware_version = self.get_oled_firmware_version()
```

**Patched code:**
```python
try:
    self._oled_firmware_version = self.get_oled_firmware_version()
except Exception:
    self._oled_firmware_version = 0  # K240 has no OLED, skip
```

**Why:** The OLED firmware version is only used for OLED-related features that the K240 doesn't have. Silently defaulting to `0` has no effect on fan/pump control.

---

## Project Structure

```
liquidctl/
├── gui.py              # Main application
├── icon.ico            # Tray/exe icon
├── settings.json       # Auto-saved settings (created on first run)
├── dist/
│   └── MSI Coreliquid Control.exe   # Built standalone executable
└── README.md
```

---

## Acknowledgements

- [liquidctl](https://github.com/liquidctl/liquidctl) — the underlying device communication library
- [pystray](https://github.com/moses-palmer/pystray) — system tray integration
- [Pillow](https://python-pillow.org/) — icon rendering

---

## License

MIT
