# KLE Importer Pro (EasyEDA Pro Extension)

This extension imports a KLE (Keyboard Layout Editor) JSON/TXT file and **places footprints based on `SWxx` written in the KLE legends**.

If a key legend contains `SW4`, the PCB footprint with designator `SW4` is moved to that key position.  
Keys without `SWxx` are ignored.

## Features

- Parse `SW1`, `SW2`, ... from KLE legend strings
- Move PCB footprints `SW1`, `SW2`, ... to the corresponding key positions
- Optionally place diodes `D1`, `D2`, ... with an offset (mm) from the switch center
- Export a simple switch plate as SVG/DXF (switch cutouts + rectangular outline)
  - Outline is a convex-hull polygon + margin offset

## Requirements

- Run it in **EasyEDA Pro PCB editor**
- Your PCB must already contain footprints `SW1`, `SW2`, ... (and optionally `D1`, `D2`, ...)
- Your KLE key legends must include `SWxx` somewhere (e.g. `"#\\n3\\n...\\nSW4"` is fine as long as it contains `SW4`)

## Installation

1. Get the released `.eext` file (e.g. `kle-importer-pro_v0.1.20.eext`)
2. EasyEDA Pro → **Extensions Manager** → **Import Extensions** → select the `.eext`

## Usage

1. Open the PCB editor in EasyEDA Pro
2. Top menu → **KLE Importer Pro** → **Open...**
3. Click **Select JSON/TXT** and choose your KLE file
4. Set pitch and diode options (optional)
5. Click **Run** (placement) or **Export Switch Plate** (SVG/DXF)

### How to write SW numbers in KLE (important)

- Put `SW<number>` in each key’s legend text
- Example: a key containing `SW4` will move PCB footprint `SW4` to that key
- Keys without `SWxx` are ignored

## Troubleshooting

### “SWxx not found in legends”
- Make sure your KLE legends contain strings like `SW1`

### “PCB API not available”
- You must run it in the **PCB editor** (not schematic-only)

### “Invalid or unexpected token”
- The KLE file may not be valid JSON (or may contain extra prefix/suffix)
- The extension tries BOM stripping and array extraction; if it still fails, check the error dialog `preview:` text to find the culprit

## Development / Build

```bash
npm ci
npm run build
```

The artifact will be created at `build/dist/kle-importer-pro_v<version>.eext` (version comes from `extension.json`).

## Release (GitHub)

1. Update `extension.json` version
2. Update `package.json` + `package-lock.json` version to match
3. `npm ci` → `npm run build`
4. Create a GitHub Release and attach the `.eext` (or use GitHub Actions)

### GitHub Actions (auto release)

- Pushing a tag like `v0.1.20` triggers CI to build and attach the `.eext` to the Release.
- The tag version must match `extension.json` and `package.json`.

## License

Apache-2.0 (see `LICENSE`)

