---
title: Working with Ranges and Cells
track: beginner
order: 3
summary: Reference ranges directly instead of Select/Activate.
draft: false
---

## Goal

Read and write cell values using object references, the pattern used in maintainable workplace macros.

## Example

```vb
Option Explicit

Sub WriteStatusColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long

    Set ws = ThisWorkbook.Worksheets("Report")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        If ws.Cells(rowIndex, "C").Value = "" Then
            ws.Cells(rowIndex, "D").Value = "Missing"
        Else
            ws.Cells(rowIndex, "D").Value = "OK"
        End If
    Next rowIndex
End Sub
```

## Common mistakes

- Using `Range("A1").Select` followed by `ActiveCell`—slow and brittle.
- Hard-coding `1000` as last row instead of finding the last used cell.

## Next up

Loop through worksheets when the same validation runs on multiple tabs.
