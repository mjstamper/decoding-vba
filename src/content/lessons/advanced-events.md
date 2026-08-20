---
title: Workbook and Worksheet Events
track: advanced
order: 2
summary: Trigger validation and logging when users change key cells or open the file.
draft: false
---

## Goal

Automate guardrails with `Workbook_Open`, `Worksheet_Change`, and related events.

## Example

In the `Report` sheet module:

```vb
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("C:C")) Is Nothing Then
        Application.EnableEvents = False
        ' Re-run lightweight validation for changed rows
        Application.EnableEvents = True
    End If
End Sub
```

Always toggle `Application.EnableEvents` when event handlers write back to the sheet to avoid infinite loops.

## Next up

Optimize array vs Range access for large datasets.
