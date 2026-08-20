---
title: Find, Copy, and Filter Data
track: intermediate
order: 5
summary: Automate the copy-paste-filter routines that eat hours every close.
draft: false
---

## Goal

Locate values, copy subsets to staging tabs, and apply AutoFilter programmatically.

## Example

```vb
Option Explicit

Sub CopyRegionEastRows()
    Dim ws As Worksheet
    Dim staging As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    Set staging = ThisWorkbook.Worksheets("East")

    ws.Rows(1).Copy Destination:=staging.Rows(1)
    ws.Range("A1").CurrentRegion.AutoFilter Field:=3, Criteria1:="East"

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ws.Range("A2:A" & lastRow).EntireRow.SpecialCells(xlCellTypeVisible).Copy _
        Destination:=staging.Cells(staging.Rows.Count, "A").End(xlUp).Offset(1, 0)

    ws.AutoFilterMode = False
End Sub
```

## Next up

Parse and clean text with VBA string functions.
