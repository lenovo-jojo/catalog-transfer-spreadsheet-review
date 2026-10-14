# spreadsheet-review

current iteration: **catalog transfer review**. other teams can adapt this repository for their own spreadsheet workflows; configurable team profiles are future work.

a desktop workflow tool for extracting unique claimant records from catalog-transfer reports. originally developed with ai assistance in august 2026; source recovered and prepared for this portfolio in october 2026.

## the problem

a transfer report can repeat the same claimant across many catalog rows. the task is to extract claimant details only where the report marks the transfer answer yes, then remove duplicate administrator/publisher records.

## what it does

- reads the first worksheet of an `.xlsx` report.
- accepts `YES`, `Yes`, `Y`, and `(Y)` in column e, ignoring surrounding whitespace and case.
- checks all work answers in column e: all yes means **true full**, any no means **true partial**, and blanks are ignored. one workbook represents one application. unrecognized nonblank answers are flagged for review; they cannot establish a full transfer.
- extracts columns r, q, n, m in this order: **admin, admin p#, publisher, publisher p#**. member fields are excluded from both clipboard and csv output.
- preserves the first occurrence of each distinct six-field record.
- writes a csv alongside the report and copies a tab-separated table to the clipboard.
- reports unsupported files and reports with no matching rows.

## operational context and personal outcome

i am one of three people on a team processing publishing catalog transfers globally. hubspot automations and strict triage helped the team clear a backlog spanning several years; this tool helped us sustain that progress by simplifying repeated claimant extraction.

in personal workflow testing, i estimated that my daily processing capacity increased from roughly **5 to 20 transfers when needed**. this is a self-reported estimate of my capacity, not a controlled benchmark or a measured team-wide result. the tool was one part of a broader workflow improvement.

## user workflow and design decisions

double-click the windows installer after extracting the download. it installs into the current user app folder and automatically creates send to and desktop shortcuts. right-click one report → send to → mlc claimant drop (use show more options first if needed). the windows send to workflow launches the tool with one report. it copies a four-column table with headers, saves the same table as csv, and shows the transfer classification in the confirmation popup. no new browser interface is required for this workflow.

the classification examines all work rows before filtering yes rows: a work marked no establishes a partial transfer regardless of what the application claimed. deduplication uses the four exported fields, so differences in member information do not duplicate the same administrator/publisher record. the csv is written before updating the clipboard, so a failed save does not replace the user's clipboard.

## implementations

the [macos source](src/claimants.swift) reads workbook xml using foundationxml. it resolves the first worksheet through workbook relationships, supports shared and inline strings, and handles csv quoting. the [windows version](windows/README.txt) uses desktop excel through powershell com automation.

```sh
swiftc src/claimants.swift -o claimants
./claimants /path/to/report.xlsx
```

requires macos, swift command-line tools, and `/usr/bin/ditto`. windows requires desktop excel; its runtime has not been verified on windows during this portfolio pass. the optional `CLAIMANTS_NO_CLIPBOARD=1` environment variable disables clipboard writes for automated checks.

## verification

```sh
swiftc src/claimants.swift -o /tmp/claimants
CLAIMANTS_BINARY=/tmp/claimants python3 -m unittest discover -s tests -v
```

tests use invented rows and check filtering, deduplication, worksheet relationship resolution, csv escaping, no-match errors, and unsupported extensions. no real catalog report, claimant data, or original data screenshot is included.

## my role and ai assistance

i brought the catalog-transfer workflow and filtering requirements, iterated on the tool with ai assistance in codex, and worked toward a desktop drop-file workflow. ai helped generate implementation and packaging code. this portfolio documents that collaboration rather than claiming all code was handwritten.

## sprints and epics

| work | status | scope and evidence |
| --- | --- | --- |
| [sprint 1 — portfolio foundation](https://github.com/lenovo-jojo/portfolio/milestone/1) | open, october 12–18, 2026 | this tool's [synthetic verification task](https://github.com/lenovo-jojo/portfolio/issues/3) is completed; the overall sprint remains open |
| [epic — publish evidence-backed portfolio projects](https://github.com/lenovo-jojo/portfolio/issues/1) | completed | recovered source, documented behavior, and automated verification |
| epic — validate partial and full transfers | implemented; windows runtime verification pending | confirmed work-level rules; classification before yes filtering; four-column export; synthetic classification and deduplication checks |

business rules and output format were confirmed with the workflow owner. windows popup, clipboard, and send to execution still need verification on windows with desktop excel.

## limits and next work

the input schema uses fixed columns; a different report layout needs validation or configurable mapping. formula values use stored workbook values rather than recalculation. large-file behavior, untrusted workbook extraction, clipboard tab/newline handling, spreadsheet formula-injection protection, and windows compatibility deserve further work. no measured time-savings claim is made.
