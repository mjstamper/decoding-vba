---
title: MsgBox and Basic User Feedback
track: beginner
order: 7
summary: Confirm actions and report macro results to the person running the workbook.
draft: false
---

## Goal

Give clear success and error messages without blocking automation unnecessarily.

## Example

```vb
Option Explicit

Sub RunWithConfirmation()
    Dim response As VbMsgBoxResult

    response = MsgBox("Export the current report?", vbYesNo + vbQuestion, "Decoding VBA")
    If response = vbYes Then
        ' Call export routine here
        MsgBox "Export complete.", vbInformation, "Decoding VBA"
    End If
End Sub
```

## Next up

Build your first end-to-end workplace macro.
