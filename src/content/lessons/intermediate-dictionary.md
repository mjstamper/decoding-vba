---
title: Dictionary for Lookups and Deduplication
track: intermediate
order: 2
summary: Replace nested loops with O(1) key lookups using Scripting.Dictionary.
draft: false
---

## Goal

Match IDs, dedupe lists, and count occurrences without scanning the whole sheet repeatedly.

## Setup

Add a reference to **Microsoft Scripting Runtime**, or use late binding:

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

## Example

```vb
Option Explicit

Sub ListUniqueOwners()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim dict As Object
    Dim owner As String

    Set ws = ThisWorkbook.Worksheets("Report")
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For rowIndex = 2 To lastRow
        owner = Trim(CStr(ws.Cells(rowIndex, "B").Value))
        If Len(owner) > 0 Then dict(owner) = True
    Next rowIndex

    ws.Range("F2").Resize(dict.Count, 1).Value = Application.Transpose(dict.Keys)
End Sub
```

## Next up

Split reusable logic into Functions vs Subs.
