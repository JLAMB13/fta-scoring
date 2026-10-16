fortheARTIST Status Scoring

A Google Sheets tool I built for fortheARTIST, a New York nonprofit that supports artists with monthly guaranteed income, coaching, and pro bono professional help. It uses the Claude API to score written evidence about each artist (survey answers, coaching notes, reports) against the organization's five-domain career sustainability rubric.

The problem

fortheARTIST wanted to know whether its support actually helps artists build careers that last. Almost everything they collect is written: survey responses, reflections, coaching notes. Scoring that by hand for every artist at two points in the year is slow and inconsistent, and the staff who run it are not programmers.

What it does

Reads the rubric off the Rubric tab on every run, so staff can change the rubric wording without touching the code
Builds a scoring prompt from the rubric plus one artist's evidence and sends it to the Claude API
Requires 1 to 3 verbatim quotes and a short rationale for every score, so each score can be checked against the source
Validates the response (right number of domains, every score between 1 and 5) before anything is written to the sheet
Logs every run, including the raw model output, to an append-only Log tab
Imports PDFs and spreadsheets from Google Drive straight into an evidence cell
Has a manual fallback: copy the prompt, run it in Claude yourself, and paste the JSON back in
How it measures change

This workbook scores an artist's standing at pre-award and at post-award separately, then subtracts the two on the Difference tab. A companion workbook scores change directly in a single judgment. Running both on the same artist is the point: when the two methods agree, they back each other up, and when they disagree, the gap is worth looking into.

Design decisions

AI is the first pass, not the final word. Staff read the quotes and rationale beside each score and adjust anything that looks wrong before trusting it.
Correcting for positive response bias. Artists receive unconditional income and tend to write warmly about the program. The prompt tells the model to score concrete circumstances, not the tone of the writing.
High scores have to be earned. When the evidence sits between two levels, the lower score wins.
Thin evidence is flagged separately. An evidence sufficiency flag separates "scored low because the artist is struggling" from "scored low because there was nothing to go on."
Privacy. The workbook uses artist IDs only, never names. The API key is stored in the script's private properties, not in any cell.
Setup

In the workbook, open Extensions > Apps Script, paste status_scoring.gs, and save.
Under Services, add the Drive API (v2). This is only needed for PDF and spreadsheet import.
Reload the sheet, then run FTA Status > Set or update API key.
Select a filled-in row on Pre (Status) or Post (Status) and run FTA Status > Score active row.
How I built it

Started with an off-the-shelf tool and pivoted. I first piloted QualCoder, a qualitative coding program, with fortheARTIST. It wasn't a good fit for how their staff would actually use it, so I switched to a simpler custom tool built into a Google Sheets workbook.
Built the rubric from their real data. I worked through fortheARTIST's actual pre-award and post-award surveys, mid-year check-in transcripts, and application forms so the rubric would match the information they already collect.
Proposed scoring two ways. A status score at each timepoint, plus a separate change rubric, because what fortheARTIST cares about most is how far artists move during the grant year.
Anchored everything to their mission. Both rubrics follow the leadership team's own definition of a sustainable career, and I checked the structure with them before building the tooling around it.
Used Claude throughout as a thinking partner and coding assistant: drafting and tightening rubric language, and writing and debugging the Apps Script.
Note

The rubric content and all artist evidence belong to fortheARTIST and are not included here. Shared with their permission.
