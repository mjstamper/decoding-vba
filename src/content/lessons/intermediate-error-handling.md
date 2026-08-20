---
title: Error Handling That Survives Production
track: intermediate
order: 4
summary: Use On Error GoTo, clean up objects, and leave workbooks in a safe state.
featured: true
draft: false
---

## Goal

Write macros that fail gracefully: restore Excel settings, log the error, and tell the user what happened.

## When to use this

Any macro that runs on shared workbooks, scheduled tasks, or month-end files where a half-finished run is worse than no run.

## Pattern

```vb
Option Explicit

Public Sub ExportValidatedReport()
    On Error GoTo CleanFail

    Dim ws As Worksheet
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation

    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Worksheets("Report")
    If ws Is Nothing Then Err.Raise vbObjectError + 513, , "Report sheet not found."

    ' ... validation and export logic ...

CleanExit:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

CleanFail:
    MsgBox "Export failed: " & Err.Description, vbCritical, "Decoding VBA"
    Resume CleanExit
End Sub
```

## Worked example: log and re-raise

For teams that want a paper trail, write to a hidden `Log` sheet before exiting:

```vb
Option Explicit

Private Sub WriteLog(ByVal message As String)
    Dim logSheet As Worksheet
    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    If logSheet Is Nothing Then Exit Sub
    logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Now & " | " & message
End Sub

Public Sub RunMonthlyValidation()
    On Error GoTo CleanFail

    WriteLog "Validation started"
    ' ... main logic ...
    WriteLog "Validation completed"
    Exit Sub

CleanFail:
    WriteLog "Validation failed: " & Err.Description
    MsgBox Err.Description, vbCritical, "Decoding VBA"
End Sub
```

## Common mistakes

- `On Error Resume Next` around entire procedures (hides real bugs).
- Forgetting to restore `ScreenUpdating` and `Calculation` in the error path.
- Showing a generic "Error 1004" with no context for the user.

## Practice

Add a `Log` sheet to a test workbook and run the validation macro with the `Report` sheet renamed—confirm the log line is written and Excel settings are restored.

## Next up

Find, copy, and filter data without manual clicking.
