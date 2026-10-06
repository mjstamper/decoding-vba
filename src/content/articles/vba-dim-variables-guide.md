---
title: VBA Dim — Variables and Data Types Explained
summary: Declare variables with Dim, pick types for reporting macros, and use Option Explicit to catch typos.
publishedDate: 2026-09-01
tags: [dim, variables, option-explicit, beginner]
draft: true
---

<!-- OUTLINE: Target 1,700–2,100 words. Delete this comment when publishing. Set draft: false. -->

## What Dim does

<!-- WRITE: Declares name and type; scope — procedure vs module (light touch). -->

## Option Explicit

<!-- CODE: Option Explicit at top of module -->
<!-- WRITE: Compile error on typo — typo story (issueCount vs issueCnt). -->
<!-- LINK: /learn/beginner/beginner-option-explicit/ -->

## Types for Excel automation

<!-- WRITE: Table — Long (rows, counts), Double (amounts + rounding caveat), String, Boolean, Date -->
<!-- WRITE: Variant — when to avoid -->
<!-- WRITE: Worksheet, Workbook, Range — require Set -->

| Type | Use for |
|------|---------|
| Long | TODO |
| Double | TODO |
| String | TODO |
| Worksheet | TODO |

## Dim vs Set

<!-- CODE: Dim x As Long vs Dim ws As Worksheet + Set ws = ... -->

## Naming conventions

<!-- WRITE: camelCase variables; PascalCase public procedures — match Decoding VBA style. -->

## Const for magic values

<!-- CODE: Const SHEET_REPORT As String = "Report" -->

## Dim at top of procedure

<!-- WRITE: Readability habit in VBA. -->

## ByVal parameters (short)

<!-- WRITE: Functions like MapRegionCode(ByVal code As String) -->
<!-- LINK: /learn/beginner/beginner-if-select-case/ -->

## Worked scenario — validation variables

<!-- CODE: lastRow, rowIndex, issueCount, ws — fragments tied to first-macro pattern -->

## Common mistakes

<!-- WRITE: Implicit Variant everywhere. -->
<!-- WRITE: Set x = 5 type errors. -->
<!-- WRITE: No Option Explicit → silent typos. -->

## Next steps

<!-- LINK: /learn/beginner/beginner-option-explicit/ -->
<!-- LINK: vba-ranges-and-cells-guide when live -->
<!-- LINK: /learn/intermediate/intermediate-error-handling/ -->
