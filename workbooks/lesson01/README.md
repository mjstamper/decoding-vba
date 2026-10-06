# Lesson 1 workbook — VBE tour

## Download (built file)

After running the build script, the practice file is published at:

- Local path: [public/downloads/DecodingVBA-Lesson01-VBE-Tour.xlsm](../public/downloads/DecodingVBA-Lesson01-VBE-Tour.xlsm)
- Live URL: `https://decodingvba.com/downloads/DecodingVBA-Lesson01-VBE-Tour.xlsm`

## Build on Windows (requires Excel)

From the repo root:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build-lesson01-vbe-workbook.ps1
```

If Excel blocks access to the VB project, enable **Trust access to the VBA project object model** in:

**File > Options > Trust Center > Trust Center Settings > Macro Settings**

## Manual rebuild

1. Create a new `.xlsm` workbook.
2. Rename Sheet1 to `Start Here` and paste instructions from the build script or the live sheet.
3. Save as `public/downloads/DecodingVBA-Lesson01-VBE-Tour.xlsm`

## Module contents

- `HelloDecodingVBA` — first run with F5
- `ShowActiveSheetName` — optional second macro to run after editing MsgBox text
