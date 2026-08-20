---
title: Loop Through Worksheets
track: beginner
order: 4
summary: Apply the same macro logic across every sheet in a workbook.
draft: false
---

## Goal

Run a validation or formatting routine on multiple worksheets in one pass.

## Example

```vb
Option Explicit

Sub ClearFilterOnAllSheets()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    Next ws
End Sub
```

## Next up

Branch logic with If and Select Case for report rules.
