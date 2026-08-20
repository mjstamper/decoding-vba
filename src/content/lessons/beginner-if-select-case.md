---
title: If and Select Case Statements
track: beginner
order: 5
summary: Encode business rules with readable conditional logic.
draft: false
---

## Goal

Map status codes, thresholds, and categories using `If` and `Select Case`.

## Example

```vb
Option Explicit

Function MapRegionCode(ByVal code As String) As String
    Select Case UCase(Trim(code))
        Case "E": MapRegionCode = "East"
        Case "W": MapRegionCode = "West"
        Case "C": MapRegionCode = "Central"
        Case Else: MapRegionCode = "Unknown"
    End Select
End Function
```

## Next up

Repeat actions with For and For Each loops.
