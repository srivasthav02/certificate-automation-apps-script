# Student Participation Certificate Automation

## Overview
Generates and emails participation certificates automatically when a student submits a Google Form, removing manual editing work.

## Features
- Fills a Google Doc template ({{Name}}, {{Event}}, {{Date}})
- Exports each certificate as a PDF to a Drive folder
- Emails the PDF to the student and marks the row as `Sent`
- Batch mode and on-form-submit mode

## Tech Stack
Google Apps Script, Google Forms, Sheets, Docs, Drive

## Setup
1. Create a Google Form (Name, Email, Event) linked to a Sheet; add a `Status` column (E).
2. Create a certificate Google Doc with `{{Name}}`, `{{Event}}`, `{{Date}}`.
3. In the Sheet: Extensions > Apps Script, paste `Code.gs`, and set the template doc ID and folder ID in `CONFIG`.
4. Add an installable trigger: `onFormSubmit` > From spreadsheet > On form submit.
5. Authorize when prompted.

## Author
Beeravelli Srivasthav
