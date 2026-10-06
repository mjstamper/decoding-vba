---
title: VBA MsgBox — A Complete Guide
summary: MsgBox buttons, icons, return values, and when to confirm actions vs log silently in production macros.
publishedDate: 2026-09-01
tags: [msgbox, beginner, user-interface]
draft: true
---

<!-- OUTLINE: Target 1,600–2,000 words. Delete this comment when publishing. Set draft: false. -->

## What MsgBox does

<!-- WRITE: Modal dialog; blocks until click; returns button choice. -->
<!-- WRITE: Workplace uses — confirm delete/export, report counts, simple errors. -->

## Basic syntax

<!-- WRITE: MsgBox prompt, [buttons], [title], [helpfile, context] -->
<!-- CODE: minimal Hello MsgBox -->

```vb
' TODO: basic MsgBox
```

## Return value

<!-- WRITE: Dim response As VbMsgBoxResult -->
<!-- WRITE: Table or list — vbYes, vbNo, vbOK, vbCancel, vbAbort, vbRetry, vbIgnore -->
<!-- CODE: If response = vbYes Then -->

## Button combinations

<!-- WRITE: vbOKOnly, vbOKCancel, vbYesNo, vbYesNoCancel, vbRetryCancel -->
<!-- WRITE: When Yes/No vs OK/Cancel for "Run export?" -->

## Icons and severity

<!-- WRITE: vbCritical, vbQuestion, vbExclamation, vbInformation -->
<!-- CODE: vbYesNo + vbQuestion -->
<!-- WRITE: Don't use Critical for normal questions. -->

## Default button (optional)

<!-- WRITE: vbDefaultButton1/2/3 — keep short. -->

## Title bar text

<!-- WRITE: Consistent title — macro name or Decoding VBA. -->

## InputBox sidebar

<!-- WRITE: Simple parameter entry; 3-line example. -->
<!-- WRITE: Heavy input → UserForms lesson later. -->
<!-- LINK: /learn/intermediate/intermediate-userforms/ -->

## When not to use MsgBox in production

<!-- WRITE: Unattended runs, batch files, no user at keyboard. -->
<!-- WRITE: Use log sheet instead. -->
<!-- LINK: /learn/intermediate/intermediate-error-handling/ -->

## Worked scenario — confirm then report

<!-- CODE: full Sub — confirm export → placeholder export → success MsgBox with count -->

## Common mistakes

<!-- WRITE: Ignoring return value. -->
<!-- WRITE: Comparing to "Yes" string instead of vbYes. -->
<!-- WRITE: MsgBox inside loop (one summary at end). -->

## Next steps

<!-- LINK: /learn/beginner/beginner-msgbox/ -->
<!-- LINK: article vba-if-statement-guide when live -->
<!-- LINK: /learn/beginner/beginner-first-macro/ -->
