---
title: VBA Select Case — A Complete Guide
summary: Replace long ElseIf chains with Select Case for status codes, categories, and readable report rules.
publishedDate: 2026-10-22
tags: [select-case, if-statement, beginner, validation]
draft: true
---

<!-- OUTLINE: Target 1,600–2,000 words. Delete this comment when publishing. Set draft: false. -->

## When Select Case beats If…ElseIf

<!-- WRITE: One expression, many discrete outcomes (status codes, regions, letter grades). -->
<!-- WRITE: When If chain gets hard to scan — refactor signal. -->
<!-- LINK: /articles/vba-if-statement-guide/ -->

## Basic Select Case syntax

<!-- CODE: Case "A", Case "B", Case Else -->

```vb
' TODO: Select Case UCase(Trim(ws.Cells(r, "C").Value))
```

## Case lists and ranges

<!-- WRITE: Case 1, 2, 3 vs Case 1 To 10 vs Case Is > 100. -->
<!-- CODE: numeric buckets for amount bands -->

## Case with strings and Trim

<!-- WRITE: exported data has spaces; normalize with Trim/UCase. -->
<!-- CODE: Case "OK", "REVIEW", "HOLD" with Case Else -->

## Select Case inside a loop

<!-- WRITE: status column D from code in column C. -->
<!-- CODE: 5–8 line loop with Select Case -->
<!-- LINK: /learn/beginner/beginner-first-macro/ -->

## Select Case vs If for mixed conditions

<!-- WRITE: Select Case = one variant expression; use If when combining And/Or across columns. -->

## Common mistakes

<!-- WRITE: Case Else forgotten — silent fall-through to nothing. -->
<!-- WRITE: Case Is = vs Case = for ranges. -->
<!-- WRITE: Comparing dates stored as serial numbers vs strings. -->

## Next steps

<!-- LINK: /learn/beginner/beginner-if-select-case/ -->
<!-- LINK: /articles/vba-if-statement-guide/ -->
<!-- LINK: /learn/beginner/beginner-loops/ -->
