# Lesson outline — Option Explicit and Variable Declarations

**Publish into:** [src/content/lessons/beginner-option-explicit.md](../../src/content/lessons/beginner-option-explicit.md)

**Frontmatter (keep as-is):**

```yaml
title: Option Explicit and Variable Declarations
track: beginner
order: 2
summary: Require variable declarations and avoid silent bugs from typos.
draft: false
```

**Target:** 900–1,200 words | **Code blocks:** 2–3 | **Template:** [beginner-first-macro.md](../../src/content/lessons/beginner-first-macro.md)

---

## Goal

<!-- WRITE: Option Explicit on every module; typed Dim; compile-time typo catch. -->

## When to use this

<!-- WRITE: Always on standard modules; especially shared team workbooks. -->

## Enable Option Explicit

<!-- CODE: Option Explicit line 1 -->
<!-- WRITE: VBE Tools → Options → Require Variable Declaration for new modules only -->

## The typo story

<!-- WRITE: Without — issueCount vs issueCnt → Variant 0, wrong results. -->
<!-- WRITE: With — compile error Variable not defined. -->

## Dim walkthrough

<!-- CODE: CountFilledRows — expand with prose line-by-line (ws, lastRow, End xlUp). -->

```vb
Option Explicit

Sub CountFilledRows()
    ' TODO: full example
End Sub
```

## Types you need first

<!-- WRITE: mini-table Long, String, Worksheet + Set; Double for currency -->

## Worked example

<!-- WRITE: optional Const for sheet name string; same CountFilledRows or variant -->

## What to notice

<!-- WRITE: Option Explicit top; Dim before use; Set for objects -->

## Common mistakes

<!-- WRITE: Option Explicit in one module only -->
<!-- WRITE: String vs empty cell / Null — CStr -->
<!-- KEEP: Variant when type known; omitting Option Explicit -->

## Practice

<!-- WRITE: Introduce typo on purpose; fix compile error -->
<!-- WRITE: Add loop row 2 to lastRow, MsgBox count -->

## Next up

<!-- LINK: /learn/beginner/beginner-ranges-and-cells/ -->
<!-- LINK: vba-dim-variables-guide article when live -->
