---
title: VBA For Loop - A Practical Guide
summary: Choose For, For Each, and Do While loops for reporting macros with copy-paste examples.
publishedDate: 2026-08-19
tags: [loops, beginner, reporting]
draft: false
---

## Why loops matter in workplace VBA

Most reporting macros repeat the same action for each row, sheet, or file. Loops are how you avoid recording the same clicks fifty times.

## For...Next (fixed iterations)

Use when you know the start and end row:

```vb
Option Explicit

Sub NumberRows()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    For i = 2 To 100
        ws.Cells(i, "A").Value = i - 1
    Next i
End Sub
```

Prefer finding the last row dynamically:

```vb
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
For i = 2 To lastRow
```

## For Each (objects and ranges)

Use when iterating worksheets, workbooks, or each cell in a range:

```vb
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Debug.Print ws.Name
Next ws
```

## Do While / Do Until

Use when the exit condition is logical, not a row count:

```vb
Dim rowIndex As Long
rowIndex = 2

Do While Len(ws.Cells(rowIndex, "A").Value) > 0
    ' process row
    rowIndex = rowIndex + 1
Loop
```

## Performance note

For thousands of rows, read into an array, loop the array, write back once. See the Intermediate track lesson on arrays.

## Common mistakes

- Infinite loops when the exit condition never becomes true.
- Using `For Each` on a range when you need the row index—use `For i = ...` instead.
- Nested loops calling `Cells` repeatedly—batch with arrays when possible.

## Next steps

- [Beginner: For and For Each Loops](/learn/beginner/beginner-loops/)
- [Intermediate: VBA Arrays for Bulk Data](/learn/intermediate/intermediate-arrays/)
