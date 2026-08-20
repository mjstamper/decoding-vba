---
title: Introduction to UserForms
track: intermediate
order: 7
summary: Build a small form for parameters instead of hard-coding values in macros.
draft: false
---

## Goal

Let users pick a date range, file path, or region without editing VBA.

## Outline

1. Insert a UserForm from the VBE (**Insert > UserForm**).
2. Add a Label, TextBox, and CommandButton.
3. Handle the button click to store values in module-level variables or custom properties.

## Starter pattern

```vb
' Module: modParameters
Option Explicit

Public ReportStartDate As Date
Public ReportEndDate As Date

Public Sub ShowReportForm()
    frmReportDates.Show
End Sub
```

In the form's OK button:

```vb
Private Sub cmdRun_Click()
    ReportStartDate = CDate(txtStart.Value)
    ReportEndDate = CDate(txtEnd.Value)
    Me.Hide
End Sub
```

## Next up

Advanced track: encapsulate behavior in class modules.
