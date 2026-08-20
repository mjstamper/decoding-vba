---
title: Class Modules for Maintainable Tools
track: advanced
order: 1
summary: Model report rows and validators as classes instead of scattered public variables.
featured: true
draft: false
---

## Goal

Replace module-level globals with objects that carry data and behavior—easier to test, extend, and hand off.

## When to use this

When a macro grows past a few hundred lines, or when multiple procedures share the same row-level rules.

## Worked example

Create a class module named `cReportRow`:

```vb
' Class module: cReportRow
Option Explicit

Private pOwner As String
Private pAmount As Double
Private pStatus As String

Public Property Get Owner() As String
    Owner = pOwner
End Property

Public Property Let Owner(ByVal value As String)
    pOwner = Trim(value)
End Property

Public Property Get Amount() As Double
    Amount = pAmount
End Property

Public Property Let Amount(ByVal value As Double)
    pAmount = value
End Property

Public Property Get Status() As String
    Status = pStatus
End Property

Public Sub Validate()
    If Len(pOwner) = 0 Then
        pStatus = "Missing owner"
    ElseIf pAmount < 0 Then
        pStatus = "Negative amount"
    Else
        pStatus = "OK"
    End If
End Sub
```

Use it from a standard module:

```vb
Option Explicit

Sub ValidateWithClasses()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim reportRow As cReportRow

    Set ws = ThisWorkbook.Worksheets("Report")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For rowIndex = 2 To lastRow
        Set reportRow = New cReportRow
        reportRow.Owner = CStr(ws.Cells(rowIndex, "B").Value)
        reportRow.Amount = CDbl(ws.Cells(rowIndex, "C").Value)
        reportRow.Validate
        ws.Cells(rowIndex, "D").Value = reportRow.Status
    Next rowIndex
End Sub
```

## Why this scales

- Validation rules live on the object, not duplicated across three Subs.
- New fields (e.g., `Region`) get one Property and one update to `Validate`.
- You can later add a `cReportCollection` class to batch-load from arrays.

## Common mistakes

- Creating a class but still using public module variables everywhere.
- Instantiating with `Dim x As New cReportRow` inside tight loops when reuse patterns matter—prefer explicit `Set x = New cReportRow` for clarity.

## Practice

Add a `Region` property. Fail validation when Region is blank but Amount is over 10,000.

## Next up

React to workbook and worksheet events instead of manual button clicks only.
