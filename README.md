# catalog transfer spreadsheet review

A desktop workflow tool for extracting unique claimant records from catalog-transfer reports. Originally developed with AI assistance in August 2026; source recovered and prepared for this portfolio in October 2026.

## The problem

A transfer report can repeat the same claimant across many catalog rows. The task is to extract claimant details only where the report marks the transfer answer Yes, then remove duplicate administrator/publisher records.

## What it does

- Reads the first worksheet of an `.xlsx` report.
- Accepts `YES`, `Yes`, `Y`, and `(Y)` in column E, ignoring surrounding whitespace and case.
- Checks all work answers in column E: all Yes means **True Full**, any No means **True Partial**, and blanks are ignored. One workbook represents one application. Unrecognized nonblank answers are flagged for review; they cannot establish a full transfer.
- Extracts columns R, Q, N, M in this order: **Admin, Admin P#, Publisher, Publisher P#**. Member fields are excluded from both clipboard and CSV output.
- Preserves the first occurrence of each distinct six-field record.
- Writes a CSV alongside the report and copies a tab-separated table to the clipboard.
- Reports unsupported files and reports with no matching rows.

## Operational context and personal outcome

I am one of three people on a team processing publishing catalog transfers globally. HubSpot automations and strict triage helped the team clear a backlog spanning several years; this tool helped us sustain that progress by simplifying repeated claimant extraction.

In personal workflow testing, I estimated that my daily processing capacity increased from roughly **5 to 20 transfers when needed**. This is a self-reported estimate of my capacity, not a controlled benchmark or a measured team-wide result. The tool was one part of a broader workflow improvement.

## User workflow and design decisions

Double-click the Windows installer after extracting the download. It installs into the current user app folder and automatically creates Send To and desktop shortcuts. Right-click one report → Send to → MLC Claimant Drop (use Show more options first if needed). The Windows Send To workflow launches the tool with one report. It copies a four-column table with headers, saves the same table as CSV, and shows the transfer classification in the confirmation popup. No new browser interface is required for this workflow.

The classification examines all work rows before filtering Yes rows: a work marked No establishes a partial transfer regardless of what the application claimed. Deduplication uses the four exported fields, so differences in member information do not duplicate the same administrator/publisher record. The CSV is written before updating the clipboard, so a failed save does not replace the user's clipboard.

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
| Epic — Validate partial and full transfers | Implemented; Windows runtime verification pending | Confirmed work-level rules; classification before Yes filtering; four-column export; synthetic classification and deduplication checks |

Business rules and output format were confirmed with the workflow owner. Windows popup, clipboard, and Send To execution still need verification on Windows with desktop Excel.

## Limits and next work

The input schema uses fixed columns; a different report layout needs validation or configurable mapping. Formula values use stored workbook values rather than recalculation. Large-file behavior, untrusted workbook extraction, clipboard tab/newline handling, spreadsheet formula-injection protection, and Windows compatibility deserve further work. No measured time-savings claim is made.
