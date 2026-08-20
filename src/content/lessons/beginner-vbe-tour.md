---
title: Tour the Visual Basic Editor
track: beginner
order: 1
summary: Open the VBE, understand the Project Explorer, and run your first Sub from the Immediate Window.
draft: false
---

## Goal

Get comfortable inside the Visual Basic Editor (VBE) so you can write and run code instead of recording fragile macros.

## When to use this

Use the VBE whenever a recorded macro needs editing, or when you want repeatable automation that survives workbook changes.

## Steps

1. Press `Alt + F11` to open the VBE.
2. In the **Project Explorer**, find `VBAProject (YourWorkbook.xlsm)`.
3. Insert a **Module**: right-click the project, choose **Insert > Module**.
4. Paste a simple Sub and press `F5` to run it.

## Example

```vb
Sub HelloDecodingVBA()
    MsgBox "VBE is open. You are ready to write code."
End Sub
```

## Common mistakes

- Saving as `.xlsx` instead of `.xlsm` (macros are stripped).
- Writing code in the worksheet module when a standard module is clearer for reusable Subs.

## Next up

Turn on `Option Explicit` so VBA catches typos before they break your macros.
