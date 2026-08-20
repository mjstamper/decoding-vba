---
title: Application Structure and Modules
track: advanced
order: 4
summary: Organize code into modUtilities, modExport, and class layers for team handoffs.
draft: false
---

## Goal

Layout a workbook project so the next analyst knows where to look.

## Suggested structure

| Module | Responsibility |
| --- | --- |
| `modMain` | Public entry Subs wired to buttons |
| `modValidation` | Rules and checks |
| `modExport` | File output and formatting |
| `modConstants` | Sheet names, column indexes |
| `cReportRow` | Row-level object (class) |

## Next up

Lightweight testing patterns before you ship to production.
