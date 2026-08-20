---
title: Your First Workplace Macro
track: beginner
order: 8
summary: Build a macro that validates a report tab and writes a summary status column.
featured: true
draft: false
---

## Goal

Combine variables, loops, and conditionals into one macro you could run before sending a report.

## When to use this

Run this pattern when a recurring export needs a quick validation pass: blank checks, sign checks, or required field checks.

## Worked example

Imagine a `Report` sheet with amounts in column C and owners in column B. You want column D to flag rows that fail basic rules.

```vb
Option Explicit

Sub ValidateReportTab()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim issueCount As Long

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Report")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        ws.Cells(rowIndex, "D").ClearContents

        If Len(ws.Cells(rowIndex, "B").Value) = 0 Then
            ws.Cells(rowIndex, "D").Value = "Missing owner"
            issueCount = issueCount + 1
        ElseIf Not IsNumeric(ws.Cells(rowIndex, "C").Value) Then
            ws.Cells(rowIndex, "D").Value = "Invalid amount"
            issueCount = issueCount + 1
        ElseIf ws.Cells(rowIndex, "C").Value < 0 Then
            ws.Cells(rowIndex, "D").Value = "Negative amount"
            issueCount = issueCount + 1
        Else
            ws.Cells(rowIndex, "D").Value = "OK"
        End If
    Next rowIndex

    Application.ScreenUpdating = True
    MsgBox "Validation complete. Issues found: " & issueCount, vbInformation, "Decoding VBA"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Decoding VBA"
End Sub
```

## What to notice

- `Option Explicit` and typed variables (`Long`, `Worksheet`).
- No `Select` or `Activate`—the code talks directly to `ws`.
- `ScreenUpdating = False` during the loop for speed on large tabs.
- `On Error GoTo` so a bad sheet name does not leave Excel in a half-updated state.

## Common mistakes

- Validating the active sheet instead of a named tab (breaks when someone clicks elsewhere).
- Writing "OK" only on the last row because the loop variable was wrong.
- Forgetting to re-enable `ScreenUpdating` in the error path.

## Practice

Add a rule: if column C is greater than 1,000,000, flag `"Review threshold"`. Then run the macro on a copy of a real workbook.

## Next up

Move to the Intermediate track and load data into arrays for faster processing.
