# JobAppTracker

Google Apps Script job tracker that reads emails from a Gmail label, extracts job status updates with OpenAI, and syncs results into a Google Sheet.

## Disclaimer

This project was built with assistance from OpenAI Codex.

## What this script does

- Reads threads from the Gmail label `Jobs/Inbound`
- Uses the first message id as a stable dedupe key (`thread_id`)
- Extracts `company`, `role_title`, and `status` from email content using OpenAI
- Falls back to heuristic parsing if the LLM call fails
- Inserts new rows and updates existing rows in the `Applications` tab
- Supports status precedence to avoid overwriting some manual workflow statuses

## Prerequisites

- A Google account with Gmail and Google Sheets access
- An OpenAI API key
- A Google Sheet to use as the tracker
- Google Apps Script access

## 1. Create the Google Sheet

1. Create a new Google Sheet.
2. Rename the tab to `Applications`.
3. Add this exact header row in row 1:

```text
thread_id | company | role_title | status | applied_date | status_last_updated | Last Follow up | notes
```

Header names must match exactly, including capitalization and spacing.

## 2. Create the Gmail label

1. In Gmail, create a label named `Jobs/Inbound`.
2. Add job related emails to this label manually or with Gmail filters.

## 3. Add the script in Apps Script

1. Open the sheet.
2. Go to `Extensions` -> `Apps Script`.
3. Create or open `Code.gs`.
4. Paste the full script code for this project.
5. Save.

## 4. Set Script Properties

In Apps Script:

1. Go to `Project Settings` -> `Script properties`.
2. Add:
   - `OPENAI_API_KEY` = `sk-...`
   - `OPENAI_MODEL` = `gpt-4.1-mini` (optional, defaults to this value)

## 5. Run a test

1. In Apps Script, run `testOneJobEmailExtraction_()` once.
2. Review logs in `View` -> `Logs` to confirm extraction works.
3. Run `syncJobEmails_Agentic()` manually once.
4. Confirm rows appear in the `Applications` sheet.

## 6. Authorize permissions

On first run, Google will ask for permissions:

- Gmail read access
- Spreadsheet read/write access
- External request access (`UrlFetchApp`)
- Trigger management (if you create scheduled triggers)

Approve all required scopes.

## 7. Enable automatic sync

1. Run `createJobSyncTrigger_10min()` once to create a time based trigger.
2. The sync will run every 10 minutes.

## 8. Local Git setup for this repo

This folder is initialized as a Git repository and linked to:

```text
https://github.com/AmanBhardwaj25/JobAppTracker.git
```

Use these commands to publish:

```bash
git add .
git commit -m "Initial JobAppTracker setup"
git push -u origin main
```

## Operational notes

- `maxThreadsPerRun` and `llmDelayMsBetweenCalls` should be tuned to your OpenAI rate limits.
- `status_last_updated` is used to skip threads when there is no newer message.
- `thread_id` stores the first message id and is the primary stable dedupe key.
- Keep manual follow up notes in `Last Follow up`.
