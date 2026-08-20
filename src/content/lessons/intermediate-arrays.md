---
title: VBA Arrays for Bulk Data
track: intermediate
order: 1
summary: Read ranges into arrays, process in memory, and write results back in one shot.
draft: false
---

## Goal

Speed up large updates by minimizing round-trips between VBA and the worksheet.

## Example

```vb
Option Explicit

Sub FlagLargeAmountsWithArray()
    Dim ws As Worksheet
    Dim data As Variant
    Dim results() As Variant
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Ledger")
    data = ws.Range("C2:C5000").Value
    ReDim results(1 To UBound(data, 1), 1 To 1)

    For i = 1 To UBound(data, 1)
        If IsNumeric(data(i, 1)) And data(i, 1) > 10000 Then
            results(i, 1) = "Review"
        Else
            results(i, 1) = ""
        End If
    Next i

    ws.Range("D2").Resize(UBound(results, 1), 1).Value = results
End Sub
```

## Next up

Use Scripting.Dictionary for fast lookups and deduplication.
