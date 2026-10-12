# Catalog Transfer Spreadsheet Review

A desktop workflow tool for extracting unique claimant records from catalog-transfer reports. Originally developed with AI assistance in August 2026; source recovered and prepared for this portfolio in October 2026.

## The problem

A transfer report can repeat the same claimant across many catalog rows. The task is to extract claimant details only where the report marks the transfer answer Yes, then remove duplicate six-field records.

## What it does

- Reads the first worksheet of an `.xlsx` report.
- Accepts `YES`, `Yes`, `Y`, and `(Y)` in column E, ignoring surrounding whitespace and case.
- Extracts columns K–N and Q–R: member, publisher, and administrator IDs/names.
- Preserves the first occurrence of each distinct six-field record.
- Writes a CSV alongside the report and copies a tab-separated table to the clipboard.
- Reports unsupported files and reports with no matching rows.

## Implementations

The [macOS source](src/claimants.swift) reads workbook XML using FoundationXML. It resolves the first worksheet through workbook relationships, supports shared and inline strings, and handles CSV quoting. The [Windows version](windows/README.txt) uses desktop Excel through PowerShell COM automation.

```sh
swiftc src/claimants.swift -o claimants
./claimants /path/to/report.xlsx
```

Requires macOS, Swift command-line tools, and `/usr/bin/ditto`. Windows requires desktop Excel; its runtime has not been verified on Windows during this portfolio pass. The optional `CLAIMANTS_NO_CLIPBOARD=1` environment variable disables clipboard writes for automated checks.

## Verification

```sh
swiftc src/claimants.swift -o /tmp/claimants
CLAIMANTS_BINARY=/tmp/claimants python3 -m unittest discover -s tests -v
```

Tests use invented rows and check filtering, deduplication, worksheet relationship resolution, CSV escaping, no-match errors, and unsupported extensions. No real catalog report, claimant data, or original data screenshot is included.

## My role and AI assistance

I brought the catalog-transfer workflow and filtering requirements, iterated on the tool with AI assistance in Codex, and worked toward a desktop drop-file workflow. AI helped generate implementation and packaging code. This portfolio documents that collaboration rather than claiming all code was handwritten.

## Sprints and epics

| Work | Status | Scope and evidence |
| --- | --- | --- |
| [Sprint 1 — Portfolio foundation](https://github.com/lenovo-jojo/portfolio/milestone/1) | Open, October 12–18, 2026 | This tool's [synthetic verification task](https://github.com/lenovo-jojo/portfolio/issues/3) is completed; the overall sprint remains open |
| [Epic — Publish evidence-backed portfolio projects](https://github.com/lenovo-jojo/portfolio/issues/1) | Completed | Recovered source, documented behavior, and automated verification |
| Epic — Validate partial and full transfers | Proposed; not assigned to a sprint | Confirm the business rules and required report columns; classify transfers with supporting evidence; flag insufficient or conflicting data for review; add synthetic examples |

Partial/full transfer validation is not implemented. Its first dependency is a confirmed definition of a full transfer and a partial transfer, including which report fields establish each result. The classification must run before claimant deduplication discards individual report rows.

## Limits and next work

The input schema uses fixed columns; a different report layout needs validation or configurable mapping. Formula values use stored workbook values rather than recalculation. Large-file behavior, untrusted workbook extraction, clipboard tab/newline handling, spreadsheet formula-injection protection, and Windows compatibility deserve further work. No measured time-savings claim is made.
