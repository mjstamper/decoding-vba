---
title: String Functions for Messy Data
track: intermediate
order: 6
summary: Trim, split, and normalize text fields before validation or export.
draft: false
---

## Goal

Clean inconsistent codes and descriptions imported from upstream systems.

## Example

```vb
Option Explicit

Function ParseFirstToken(ByVal text As String) As String
    Dim parts() As String
    parts = Split(Trim(text), " ")
    If UBound(parts) >= 0 Then
        ParseFirstToken = parts(0)
    End If
End Function
```

## Next up

Collect user input with a simple UserForm.
