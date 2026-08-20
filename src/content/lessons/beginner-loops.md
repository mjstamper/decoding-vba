---
title: For and For Each Loops
track: beginner
order: 6
summary: Iterate rows, columns, and collections efficiently.
draft: false
---

## Goal

Choose the right loop for rows (`For`) versus objects (`For Each`).

## Example

```vb
Option Explicit

Sub HighlightNegativeAmounts()
    Dim ws As Worksheet
    Dim amountCell As Range

    Set ws = ThisWorkbook.Worksheets("Ledger")

    For Each amountCell In ws.Range("E2:E500").Cells
        If IsNumeric(amountCell.Value) And amountCell.Value < 0 Then
            amountCell.Interior.Color = RGB(255, 230, 230)
        End If
    Next amountCell
End Sub
```

## Next up

Communicate results with MsgBox and user prompts.
