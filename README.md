# excel-formula-auditor — case study

Static analysis for Excel workbooks. Catches silent formula bugs, fragile sums, `VLOOKUP` traps, and structural risks that hide inside spreadsheets people bet payroll on.

> **This is a public case study of a private project.** Source code is not published. The intent is to describe the problem, the approach, and the impact. If you want to talk about the tool, licensing, or an audit engagement, contact me via [LinkedIn](https://www.linkedin.com/in/juan-felipe-delgado-quevedo-829b53139).

## The audit that started the project

A real payroll workbook. 124 employees. 34,500 formulas. Six sheets. Zero protection.

While reviewing it I found:

- **1,579 out of 1,619 employee entries (97.5%)** triggering a silent `ROUNDUP + TEXT(HH:MM)` bug that dropped one hour whenever the minute was nonzero.
- **792 shift exits** past a threshold hour with no upper cap, all requiring manual correction each pay cycle.
- **20 `VLOOKUP` cases** where an employee legitimately clocked in twice on the same day and the second entry was silently discarded.
- **~10,000 `VLOOKUP` calls** referencing full columns, causing the workbook to freeze on any edit.
- Sums built by chaining `+` between cell references, unprotected sheets, zero named ranges, and inconsistent categorical casing that produced orphan rows in grouped reports.

Every one of these is a class of bug that a rule-based auditor can catch in seconds. That is what this tool does.

## What the tool does

Loads an `.xlsx` workbook, runs a pluggable rule set over every cell, sheet, and workbook-level property, and produces a report with:

- Cell reference (`Sheet!A1`).
- Severity (critical, high, medium, low).
- What the rule detected.
- Suggested fix.

Two output formats: JSON for machine consumption, Markdown for humans.

## Rule categories

| Rule | Severity | What it catches |
|---|---|---|
| ROUNDUP hour entry bug | critical | The specific `ROUNDUP + TEXT("HH") + RIGHT(TEXT("HH:MM"), 2) / 60` pattern that silently drops an hour when minutes are nonzero. |
| Fragile sum by concatenation | high | Sums built by chaining `+` across cell references. Any inserted column breaks the sum. |
| Full-column VLOOKUP | high | `VLOOKUP(...,A:B,...)` that scans entire columns and destroys recalc performance on large sheets. |
| VLOOKUP duplicate-key risk | medium | Any `VLOOKUP` whose lookup range may contain duplicate keys, silently returning only the first match. |
| Unprotected sheet | medium | Sheet has no protection enabled; a misclick can wreck formulas. |
| Missing named ranges | low | Workbook declares zero named ranges. Formulas reference raw addresses that break on structural changes. |
| Categorical case mismatch | low | Categorical column contains values that differ only by case (e.g., `Prado` vs `prado`), producing orphan rows in grouped reports. |
| Orphan categorical value | low | A row's category appears nowhere else in the reference column. |

Adding a new rule is a single file plus one line to register it.

## Design principles

- **Rules are cheap and pluggable.** Each rule is a small class with an `evaluate(workbook)` method that yields structured findings. No inheritance chain deeper than one abstract base.
- **Findings are structured.** Every finding carries a location, severity, category, summary, evidence, and fix suggestion. Callers can filter, sort, or export to any downstream shape.
- **No auto-fix.** This is an auditor, not a rewriter. Spreadsheets belong to their owners. The tool tells you what is wrong; you decide.
- **Streaming friendly.** `openpyxl` loaded read-only when possible. Large workbooks (tens of thousands of formulas) audit in seconds without loading everything into memory.

## Real-world learnings baked in

- Excel's `TEXT(cell, "HH:MM")` returns a string. Adding it to a number with `+` triggers implicit conversion. `RIGHT(...,2) / 60` extracts the minutes but the `ROUNDUP` around the sum makes the entry hour jump one hour up when the minutes are nonzero. Silent, off-by-one, hard to notice at scale.
- `VLOOKUP` on `A:B` re-evaluates the entire column tree every time any cell in the workbook changes. A workbook with 10,000 of these will freeze even on modern hardware.
- Excel does not error on duplicate keys in a `VLOOKUP` lookup range. It picks the first match. When your data legitimately has multiple entries for the same key (an employee clocking in twice a day), the second one is silently lost.
- Categorical columns without a validation list drift over time: `Prado`, `prado`, `PRADO` end up as three distinct categories in every grouped report.

## Get in touch

If you have a workbook you suspect is silently wrong, or want to license the tool for internal use, contact me via [LinkedIn](https://www.linkedin.com/in/juan-felipe-delgado-quevedo-829b53139).
