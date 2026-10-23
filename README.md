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

See [HowToRun.txt](HowToRun.txt) for the step-by-step Excel setup
(importing the `.bas` module into a macro-enabled workbook) and for running
the macros via `Alt+F8`.
