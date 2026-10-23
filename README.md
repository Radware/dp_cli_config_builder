# DP CLI Config Builder

## Scope

This tool generates per-device DP CLI configuration text files from a single
template. It is an Excel macro (VBA) that merges a text template containing
`$variable` placeholders with values stored in a spreadsheet, producing one
ready-to-run configuration file per device.

Typical use case: you maintain one configuration template (e.g.
[DP Configuration Template.txt](DP%20Configuration%20Template.txt)) with
commands such as:

```
net l2-interface set 11 -ad up -ifd "DP-dirty1 -> $dirty1-interface"
system mib2-name set $device-name
services ntp server-name set $ntp-server
```

and a worksheet listing, for every device/site, the concrete value of each
`$variable`. Running the macro produces one text file per device with all
placeholders substituted.

## How it works

- **Template file**: any UTF-8 text file. Placeholders are tokens starting
  with `$` followed by letters, digits, underscores or hyphens (e.g.
  `$device-name`, `$bgp-interface`, `$dirty1-interface`). The template is
  chosen interactively when the macro runs, it is not hard-coded.
- **Variable worksheet**: a worksheet named exactly `DP Template Variables`
  in the active workbook.
  - Column A, starting at row 2, holds the variable names (case-sensitive,
    must match the `$...` tokens in the template, including the leading `$`).
  - Row 1, starting at column B, holds one column per device/site (a label
    such as the site name).
  - Each column B onward holds that device's value for every variable in
    column A.
  - Values are read as stored (`Value2`), not as Excel's displayed/formatted
    text, so values such as IDs or VLANs that must keep leading zeros should
    be stored as text in the cell.
- **Output files**: one text file per device column, UTF-8 encoded, no byte
  order mark. The file name is based on the device's `$device-name` value
  (falling back to the row-1 header) plus a column-number suffix to keep
  names unique. Existing files are never overwritten.
- Blank or missing values are substituted with the literal text `no-value`
  so that gaps in the spreadsheet are easy to spot in generated output.

## Macros

The workbook exposes three macros (see [DeviceConfigGenerator.bas](DeviceConfigGenerator.bas)):

| Macro | Purpose |
|---|---|
| `ValidateTemplateVariables` | Checks the template's `$variable` tokens against column A: reports any variable used in the template but missing from the sheet, and any duplicate variable names in column A. Does not write files. |
| `GenerateSelectedDevice` | Prompts for a single device column letter (e.g. `B`, `C`, `AA`) and generates a preview file for that device only. Tolerates missing/duplicate variables (uses `no-value` / the first matching row) and reports them. |
| `GenerateAllDeviceConfigs` | Validates all template variables first and stops if there are missing or duplicate names. If validation passes, generates one output file for every non-blank device column from B onward. |

## How to run

Excel VBA device configuration generator (Windows Excel)

SETUP
1. Open Devices Configuration Variables.xlsx in desktop Excel.
2. Save As an Excel Macro-Enabled Workbook (*.xlsm).
3. Press Alt+F11, then File > Import File; select DeviceConfigGenerator.bas.
4. Save the workbook and return to Excel.

RUN
Press Alt+F8 and choose one of these macros:
- ValidateTemplateVariables: choose the UTF-8 template; checks all template tokens against column A and reports missing/duplicate names.
- GenerateSelectedDevice: choose the device column letter (B, C, ... AA), then the template and an output folder; creates one file for that column. Blank and missing values become no-value. Duplicate names use the first matching row and are reported.
- GenerateAllDeviceConfigs: choose the template and an output folder; validates first and stops for missing/duplicate names. After validation succeeds, generates one file for each nonblank header from B onward. Blank values become no-value.

LAYOUT
Worksheet must be named DP Template Variables.
Row 1 has device/site labels; column A has case-sensitive variable names, starting at A2.
Columns B onward contain per-device values.
Variable names support letters, digits, underscores and hyphens after $.
Values are read using Value2 (not Excel's displayed number formatting). Store identifiers requiring leading zeros as text.
Output names use $device-name, falling back to the row-1 site, plus a column-number suffix for uniqueness. Existing files are never overwritten.
Templates and output use UTF-8; generated files have no byte-order mark.
No additional VBA references are required.

