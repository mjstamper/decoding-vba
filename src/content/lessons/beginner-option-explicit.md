---
title: Option Explicit and Variable Declarations
track: beginner
order: 2
summary: Require variable declarations and avoid silent bugs from typos.
draft: false
---

## Goal

Enable `Option Explicit` and declare variables with meaningful types.

## When to use this

Always. Production macros should fail fast when a variable name is misspelled.

## Example

Add this at the top of every standard module:

```vb
Option Explicit

Sub CountFilledRows()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    MsgBox "Filled rows: " & lastRow
End Sub
```

## Common mistakes

- Omitting `Option Explicit` and accidentally creating new Variant variables via typos.
- Using `Variant` when `Long`, `String`, or `Worksheet` is known.

## Next up

Read and write cells with the Range object—without Select/Activate.
