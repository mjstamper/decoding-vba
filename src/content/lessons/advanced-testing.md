---
title: Testing and Release Checklist
track: advanced
order: 5
summary: Verify macros on copies, document assumptions, and ship with a rollback plan.
draft: false
---

## Goal

Treat VBA like production code: test edge cases, document sheet names, and keep a backup path.

## Checklist

- [ ] Run on a **copy** of the production workbook first.
- [ ] Test empty sheets, single-row sheets, and max-volume sheets.
- [ ] Confirm error handlers restore `ScreenUpdating`, `Calculation`, and `EnableEvents`.
- [ ] Document required sheet names and button locations in a `README` tab.
- [ ] Version the `.xlsm` filename (`ReportTool_v1.2.xlsm`).

## Debug tip

Use the Immediate Window (`Ctrl + G`) to inspect variables:

```vb
Debug.Print reportRow.Status
```

## You have completed the Advanced track

Review the paid course for full project builds with downloadable workbooks and walkthroughs.
