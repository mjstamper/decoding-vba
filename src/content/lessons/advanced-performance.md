---
title: Performance - Arrays vs Range
track: advanced
order: 3
summary: Profile macros and choose array batch operations for large workbooks.
draft: false
---

## Goal

Keep month-end macros under a minute by minimizing Excel object calls.

## Rules of thumb

- Read once into a Variant array, process in memory, write once back.
- Turn off `ScreenUpdating` and set `Calculation = Manual` during bulk work.
- Avoid `Cells(i, j)` inside double loops when a single array dimension suffices.

## Timing snippet

```vb
Dim t As Double
t = Timer
' ... macro body ...
Debug.Print "Elapsed seconds: " & Round(Timer - t, 2)
```

## Next up

Structure a multi-module VBA application.
