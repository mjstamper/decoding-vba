---
title: VBA For Each Loop — A Complete Guide
summary: Loop over ranges, worksheets, and collections with For Each without off-by-one row mistakes.
publishedDate: 2026-10-15
tags: [loops, for-each, beginner, reporting]
draft: true
---

<!-- OUTLINE: Target 1,700–2,100 words. Delete this comment when publishing. Set draft: false. -->

## Why For Each fits workplace VBA

<!-- WRITE: When you care about "every item" not "row 2 to lastRow". -->
<!-- WRITE: Sheets in a workbook, cells in a selection, files in a folder (FileSystemObject teaser only). -->
<!-- LINK: /articles/vba-for-loop-guide/ — For vs For Each decision table. -->

## For Each…Next syntax

<!-- WRITE: Variable type must match collection (Range, Worksheet, etc.). -->
<!-- CODE: minimal For Each cell In rng -->

```vb
' TODO: For Each cell In ws.Range("B2:B100")
```

## For Each over a Range

<!-- WRITE: cell.Value read/write; avoid Select. -->
<!-- CODE: clear status column or color cells that fail a rule -->
<!-- LINK: /learn/beginner/beginner-ranges-and-cells/ -->

## For Each over Worksheets

<!-- WRITE: ThisWorkbook.Worksheets; skip VeryHidden if needed. -->
<!-- CODE: append sheet name to a summary sheet or MsgBox list -->

## For Each vs For…Next

<!-- WRITE: When index matters (row number in column A) use For. -->
<!-- WRITE: When only values matter use For Each. -->
<!-- WRITE: lastRow + For still needed for "stop at first blank" patterns. -->

## Nested loops

<!-- WRITE: Outer sheets, inner rows — keep ScreenUpdating off. -->
<!-- CODE: small nested example with Exit For when match found -->

## Exit For and Continue (Do not use GoTo for new code)

<!-- WRITE: Exit For when you found the row; note VBA has no Continue For. -->

## Common mistakes

<!-- WRITE: For Each cell without Set when variable is Object vs Range. -->
<!-- WRITE: Modifying collection while iterating (deleting rows/sheets). -->
<!-- WRITE: Assuming For Each order matches row order (usually ok for Range, document it). -->

## Next steps

<!-- LINK: /learn/beginner/beginner-loops/ -->
<!-- LINK: /articles/vba-for-loop-guide/ -->
<!-- LINK: /learn/beginner/beginner-worksheets/ -->
