---
title: Functions vs Subs
track: intermediate
order: 3
summary: Return values from Functions and keep side effects in Subs.
draft: false
---

## Goal

Write testable Functions for calculations and Subs for workbook actions.

## Example

```vb
Option Explicit

Function NormalizeAccountCode(ByVal rawCode As String) As String
    NormalizeAccountCode = UCase(Replace(Trim(rawCode), "-", ""))
End Function

Sub ApplyNormalizedCodes()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set ws = ThisWorkbook.Worksheets("COA")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        ws.Cells(rowIndex, "B").Value = NormalizeAccountCode(CStr(ws.Cells(rowIndex, "A").Value))
    Next rowIndex
End Sub
```

## Next up

Add structured error handling for production macros.
