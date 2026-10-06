---
title: VBA If Statement — A Complete Guide
summary: Use If, ElseIf, and Else for report validation, status columns, and readable business rules in Excel VBA.
publishedDate: 2026-09-01
tags: [if-statement, beginner, validation, reporting]
draft: true
---

<!-- OUTLINE: Target 1,800–2,200 words. Delete this comment when publishing. Set draft: false. -->

## Why If matters in workplace VBA

<!-- WRITE: Recorded macros hide decisions; If makes rules explicit. -->
<!-- WRITE: Typical uses — blank checks, sign checks, status column D, skip header row. -->
<!-- WRITE: Month-end tie-in — flag rows before sending a report. -->

## If…Then — one line vs block form

<!-- WRITE: One-liner when single action. -->
<!-- CODE: simple threshold on one cell -->

```vb
' TODO: one-line If example
```

<!-- WRITE: Block If…End If when multiple lines. -->
<!-- CODE: block form example -->

## If…ElseIf…Else

<!-- WRITE: First true branch wins; order matters. -->
<!-- CODE: Missing owner / Invalid amount / Negative / OK (single row or small loop) -->
<!-- LINK: /learn/beginner/beginner-first-macro/ -->

## Nesting If statements (and when not to)

<!-- WRITE: Small nested example (region + amount). -->
<!-- WRITE: Prefer ElseIf or flatten logic for readability. -->

## Logical operators — And, Or, Not

<!-- WRITE: And = all true; Or = any true; Not = invert. -->
<!-- CODE: flag if owner blank Or amount not numeric -->

## Testing for empty cells and text

<!-- WRITE: Len(Trim(...)) = 0 vs = "" vs IsEmpty — exported data quirks. -->
<!-- CODE: If Len(Trim(CStr(ws.Cells(r, "B").Value))) = 0 Then -->

## If vs IIf

<!-- WRITE: IIf for simple expressions; both branches evaluated. -->
<!-- WRITE: Prefer If block in macros. -->
<!-- CODE: one-line IIf vs If block comparison -->

## If inside loops

<!-- WRITE: Apply rule per row; 5–6 line loop. -->
<!-- LINK: /articles/vba-for-loop-guide/ -->
<!-- LINK: /learn/beginner/beginner-loops/ -->

## Common mistakes

<!-- WRITE: = vs Is for objects (brief). -->
<!-- WRITE: ElseIf after Else. -->
<!-- WRITE: Comparing Double with = (rounding). -->

## Next steps

<!-- LINK: /learn/beginner/beginner-if-select-case/ -->
<!-- LINK: /learn/beginner/beginner-first-macro/ -->
