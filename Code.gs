/***************************************
 * JOB TRACKER (AGENTIC + HYBRID DEDUPE)
 * -------------------------------------
 * REQUIRED SCRIPT PROPERTIES:
 *   OPENAI_API_KEY = sk-...
 *
 * OPTIONAL SCRIPT PROPERTIES:
 *   OPENAI_MODEL = gpt-4.1-mini   (default)
 *
 * SHEET TAB:
 *   Applications
 *
 * EXPECTED HEADERS (exact names):
 *   thread_id              <-- stores FIRST MESSAGE ID (stable key)
 *   company
 *   role_title
 *   status
 *   applied_date
 *   status_last_updated
 *   Last Follow up
 *   notes
 ***************************************/

function syncJobEmails_Agentic() {
  const CONFIG = {
    sheetName: 'Applications',
    inboundLabel: 'Jobs/Inbound',
    maxThreadsPerRun: 50,          // keep low if your RPM is low (e.g., 3 RPM)
    maxBodyChars: 5000,
    skipIfNoNewerMessage: true,
    addReviewTagInNotes: true,
    llmDelayMsBetweenCalls: 22000 // 22s helps stay under 3 RPM
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error(`Sheet "${CONFIG.sheetName}" not found.`);

  const headers = getHeaderMap_(sheet);
  ensureRequiredHeaders_(headers, [
    'thread_id',
    'company',
    'role_title',
    'status',
    'applied_date',
    'status_last_updated',
    'Last Follow up',
    'notes'
  ]);

  const inboundLabel = GmailApp.getUserLabelByName(CONFIG.inboundLabel);
  if (!inboundLabel) throw new Error(`Gmail label "${CONFIG.inboundLabel}" not found.`);

  // Build indexes once from current sheet
  const indexes = buildIndexes_(sheet, headers);

  const threads = inboundLabel.getThreads(0, CONFIG.maxThreadsPerRun);
  Logger.log(`Found ${threads.length} threads under ${CONFIG.inboundLabel}`);

  let inserted = 0;
  let updated = 0;
  let skippedNoChange = 0;
  let skippedNotJob = 0;
  let fallbackUsed = 0;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];

    try {
      const messages = thread.getMessages();
      if (!messages || messages.length === 0) continue;

      const firstMsg = messages[0];
      const latestMsg = messages[messages.length - 1];

      // Stable key (do NOT use thread.getId() for dedupe)
      const stableId = firstMsg.getId();

      const latestDate = latestMsg.getDate();
      const subject = safeString_(latestMsg.getSubject());
      const from = safeString_(latestMsg.getFrom());
      const body = safeString_(latestMsg.getPlainBody()).slice(0, CONFIG.maxBodyChars);

      // If we already have a row by stable ID and no newer message, skip
      const rowById = indexes.idIndex[stableId] || null;
      if (CONFIG.skipIfNoNewerMessage && rowById) {
        const existingStatusUpdated = sheet.getRange(rowById, headers['status_last_updated']).getValue();
        const existingDate = existingStatusUpdated instanceof Date ? existingStatusUpdated : new Date(existingStatusUpdated);

        if (existingDate && !isNaN(existingDate.getTime()) && latestDate <= existingDate) {
          skippedNoChange++;
          continue;
        }
      }

      // LLM parse with retry/backoff
      let parsed;
      let usedFallback = false;
      try {
        parsed = extractJobFieldsWithOpenAI_WithRetry_(subject, from, body);
      } catch (err) {
        Logger.log(`LLM failed for subject "${subject}". Falling back. Error: ${err}`);
        parsed = heuristicFallbackParse_(subject, from, body);
        usedFallback = true;
        fallbackUsed++;
      }

      if (!parsed.is_job_related) {
        skippedNotJob++;
        continue;
      }

      // Fill missing company with domain guess
      if (!parsed.company) {
        parsed.company = domainGuessFromFrom_(from);
      }

      // Find existing row by:
      // 1) stable message ID
      // 2) normalized company+role (manual rows)
      let existingRow = findExistingRow_(stableId, parsed, indexes);

      const normalizedStatus = normalizeStatus_(parsed.status);
      const noteText = buildNotes_(subject, parsed, usedFallback, CONFIG);

      if (!existingRow) {
        // INSERT NEW ROW
        const rowValues = new Array(sheet.getLastColumn()).fill('');

        setCellByHeader_(rowValues, headers, 'thread_id', stableId);
        setCellByHeader_(rowValues, headers, 'company', parsed.company);
        setCellByHeader_(rowValues, headers, 'role_title', parsed.role_title);
        setCellByHeader_(rowValues, headers, 'status', normalizedStatus);
        setCellByHeader_(rowValues, headers, 'applied_date', latestDate);
        setCellByHeader_(rowValues, headers, 'status_last_updated', latestDate);
        // Leave "Last Follow up" untouched
        setCellByHeader_(rowValues, headers, 'notes', noteText);

        sheet.appendRow(rowValues);
        const newRow = sheet.getLastRow();

        // update indexes in-memory for same run
        indexes.idIndex[stableId] = newRow;
        const key = makeCompanyRoleKey_(parsed.company, parsed.role_title);
        if (key && !indexes.companyRoleIndex[key]) {
          indexes.companyRoleIndex[key] = newRow;
        }

        inserted++;
      } else {
        // UPDATE EXISTING ROW (manual or prior synced)

        // Backfill stable ID into thread_id if missing (critical for future dedupe)
        const idCell = sheet.getRange(existingRow, headers['thread_id']);
        if (!idCell.getValue()) {
          idCell.setValue(stableId);
          indexes.idIndex[stableId] = existingRow;
        }

        // Only overwrite company/role if blank in sheet
        const companyCell = sheet.getRange(existingRow, headers['company']);
        const roleCell = sheet.getRange(existingRow, headers['role_title']);
        const existingCompany = safeString_(companyCell.getValue()).trim();
        const existingRole = safeString_(roleCell.getValue()).trim();

        if (!existingCompany && parsed.company) companyCell.setValue(parsed.company);
        if (!existingRole && parsed.role_title) roleCell.setValue(parsed.role_title);

        // Status precedence so we don't clobber manual statuses too aggressively
        const statusCell = sheet.getRange(existingRow, headers['status']);
        const currentStatus = safeString_(statusCell.getValue()).trim();
        const chosenStatus = chooseStatusWithPrecedence_(currentStatus, normalizedStatus);
        if (chosenStatus && chosenStatus !== currentStatus) {
          statusCell.setValue(chosenStatus);
        }

        // Update status timestamp if newer
        sheet.getRange(existingRow, headers['status_last_updated']).setValue(latestDate);

        // Preserve applied_date if already populated
        const appliedCell = sheet.getRange(existingRow, headers['applied_date']);
        if (!appliedCell.getValue()) {
          appliedCell.setValue(latestDate);
        }

        // Update notes (safe overwrite; your manual workflow notes can stay in Last Follow up)
        sheet.getRange(existingRow, headers['notes']).setValue(noteText);

        // refresh companyRole index if row was previously missing company/role
        const finalCompany = safeString_(sheet.getRange(existingRow, headers['company']).getValue()).trim();
        const finalRole = safeString_(sheet.getRange(existingRow, headers['role_title']).getValue()).trim();
        const key = makeCompanyRoleKey_(finalCompany, finalRole);
        if (key && !indexes.companyRoleIndex[key]) {
          indexes.companyRoleIndex[key] = existingRow;
        }

        updated++;
      }

      // Pace requests (important for low RPM)
      if (i < threads.length - 1) {
        Utilities.sleep(CONFIG.llmDelayMsBetweenCalls);
      }

    } catch (err) {
      Logger.log(`Error processing one thread: ${err}`);
    }
  }

  Logger.log(`Done. inserted=${inserted}, updated=${updated}, skippedNoChange=${skippedNoChange}, skippedNotJob=${skippedNotJob}, fallbackUsed=${fallbackUsed}`);
}

/**
 * LLM extraction with retry/backoff for 429 rate limits.
 */
function extractJobFieldsWithOpenAI_WithRetry_(subject, from, body) {
  const maxAttempts = 4;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return extractJobFieldsWithOpenAI_(subject, from, body);
    } catch (err) {
      const msg = String(err || '');

      const is429 = msg.includes('429');
      const rateLimitLike =
        msg.includes('rate_limit') ||
        msg.includes('Rate limit') ||
        msg.includes('Too Many Requests');

      if (!(is429 && rateLimitLike)) {
        throw err;
      }

      // Parse "Please try again in 20s" if present
      let waitMs = 25000;
      const m = msg.match(/Please try again in\s+(\d+)s/i);
      if (m) {
        const secs = parseInt(m[1], 10);
        if (!isNaN(secs)) waitMs = (secs + 2) * 1000;
      }

      // Exponential-ish backoff
      waitMs = Math.min(waitMs * attempt, 90000);

      Logger.log(`Rate limited (attempt ${attempt}/${maxAttempts}), sleeping ${waitMs}ms`);
      Utilities.sleep(waitMs);
    }
  }

  throw new Error('OpenAI rate limit retries exhausted.');
}

/**
 * Calls OpenAI Chat Completions with Structured Outputs JSON schema.
 */
function extractJobFieldsWithOpenAI_(subject, from, body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('Missing OPENAI_API_KEY in Script Properties.');

  const model = PropertiesService.getScriptProperties().getProperty('OPENAI_MODEL') || 'gpt-4.1-mini';

  const systemPrompt = [
    'You extract job application updates from emails into strict JSON.',
    'Determine if the email is job-related.',
    'If job-related, extract company, role_title, and status.',
    'Use only these statuses: Applied, Interview, Rejected, Offer, Unknown.',
    'Prefer company/role in subject/body over platform sender domains (e.g., ashbyhq, icims, lever, workday).',
    'If uncertain, set confidence=low and needs_review=true.',
    'Keep summary to one short sentence.'
  ].join(' ');

  const userPrompt = [
    `FROM: ${from}`,
    `SUBJECT: ${subject}`,
    `BODY:`,
    body
  ].join('\n\n');

  const schema = {
    name: 'job_email_extraction',
    strict: true,
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        is_job_related: { type: 'boolean' },
        company: { type: 'string' },
        role_title: { type: 'string' },
        status: { type: 'string', enum: ['Applied', 'Interview', 'Rejected', 'Offer', 'Unknown'] },
        confidence: { type: 'string', enum: ['high', 'medium', 'low'] },
        needs_review: { type: 'boolean' },
        summary: { type: 'string' }
      },
      required: ['is_job_related', 'company', 'role_title', 'status', 'confidence', 'needs_review', 'summary']
    }
  };

  const payload = {
    model: model,
    temperature: 0,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    response_format: {
      type: 'json_schema',
      json_schema: schema
    }
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload)
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`OpenAI API ${code}: ${text}`);
  }

  const json = JSON.parse(text);
  const content = json &&
    json.choices &&
    json.choices[0] &&
    json.choices[0].message &&
    json.choices[0].message.content;

  if (!content) throw new Error('No message.content in OpenAI response');

  let parsed;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    throw new Error(`Failed to parse model JSON: ${content}`);
  }

  return {
    is_job_related: !!parsed.is_job_related,
    company: safeString_(parsed.company).trim(),
    role_title: safeString_(parsed.role_title).trim(),
    status: normalizeStatus_(parsed.status),
    confidence: ['high', 'medium', 'low'].includes(parsed.confidence) ? parsed.confidence : 'low',
    needs_review: !!parsed.needs_review,
    summary: safeString_(parsed.summary).trim()
  };
}

/**
 * Fallback parser if LLM fails/quota/rate-limits.
 */
function heuristicFallbackParse_(subject, from, body) {
  const text = `${subject}\n${from}\n${body}`.toLowerCase();

  let status = 'Unknown';
  if (/(offer|congratulations)/i.test(text)) {
    status = 'Offer';
  } else if (/(not selected|unfortunately|move forward with other candidates|regret to inform)/i.test(text)) {
    status = 'Rejected';
  } else if (/(interview|screening|next steps|hiring team|schedule)/i.test(text)) {
    status = 'Interview';
  } else if (/(thank you for applying|application received|application submitted|we received your application)/i.test(text)) {
    status = 'Applied';
  }

  return {
    is_job_related: true,
    company: domainGuessFromFrom_(from),
    role_title: '',
    status: status,
    confidence: 'low',
    needs_review: true,
    summary: 'Fallback parser used.'
  };
}

/**
 * Build indexes from existing sheet:
 * - idIndex: stableId -> row
 * - companyRoleIndex: normalized "company|role" -> row (for manual rows)
 */
function buildIndexes_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const idIndex = {};
  const companyRoleIndex = {};

  if (lastRow < 2) return { idIndex, companyRoleIndex };

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;

    const stableId = safeString_(values[i][headers['thread_id'] - 1]).trim();
    const company = safeString_(values[i][headers['company'] - 1]).trim();
    const role = safeString_(values[i][headers['role_title'] - 1]).trim();

    if (stableId) idIndex[stableId] = rowNum;

    const key = makeCompanyRoleKey_(company, role);
    if (key && !companyRoleIndex[key]) {
      companyRoleIndex[key] = rowNum; // keep first if duplicates already exist
    }
  }

  return { idIndex, companyRoleIndex };
}

/**
 * Dedupe resolver:
 * 1) stable Gmail message ID
 * 2) normalized company+role for manual rows
 */
function findExistingRow_(stableId, parsed, indexes) {
  if (stableId && indexes.idIndex[stableId]) {
    return indexes.idIndex[stableId];
  }

  const key = makeCompanyRoleKey_(parsed.company, parsed.role_title);
  if (key && indexes.companyRoleIndex[key]) {
    return indexes.companyRoleIndex[key];
  }

  return null;
}

/**
 * Status precedence:
 * Don't let "Applied" overwrite stronger/manual-progress states.
 */
function chooseStatusWithPrecedence_(currentStatusRaw, newStatusRaw) {
  const current = safeString_(currentStatusRaw).trim();
  const next = normalizeStatus_(newStatusRaw);

  if (!current) return next;

  // Preserve common manual statuses if new status is just "Applied"
  const lower = current.toLowerCase();
  const preserveIfApplied = [
    'awaiting referral',
    'referred',
    'referred & submitted',
    'referral requested',
    'submitted'
  ];

  if (next === 'Applied' && preserveIfApplied.indexOf(lower) !== -1) {
    return current;
  }

  // Precedence score
  const score = {
    'Unknown': 0,
    'Applied': 1,
    'Interview': 2,
    'Rejected': 3,
    'Offer': 4
  };

  // If current is one of your manual values, treat as at least Applied
  let currentNorm = current;
  if (!score.hasOwnProperty(currentNorm)) {
    currentNorm = 'Applied';
  }

  return (score[next] >= score[currentNorm]) ? next : current;
}

/**
 * Helpers
 */
function buildNotes_(subject, parsed, usedFallback, config) {
  let note = `Subject: ${subject}`;
  if (parsed.summary) note += ` | ${parsed.summary}`;
  if (parsed.confidence) note += ` | confidence=${parsed.confidence}`;
  if (usedFallback) note += ` | parser=fallback`;
  else note += ` | parser=llm`;
  if (config.addReviewTagInNotes && parsed.needs_review) note += ` | REVIEW`;
  return note;
}

function normalizeStatus_(status) {
  const s = safeString_(status).trim();
  return ['Applied', 'Interview', 'Rejected', 'Offer', 'Unknown'].includes(s) ? s : 'Unknown';
}

function domainGuessFromFrom_(from) {
  const emailMatch = safeString_(from).match(/<([^>]+)>/);
  const senderEmail = emailMatch ? emailMatch[1] : from;
  const domainMatch = safeString_(senderEmail).match(/@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
  if (!domainMatch) return '';

  return domainMatch[1]
    .replace(/^mail\./i, '')
    .replace(/^jobs\./i, '')
    .replace(/^careers\./i, '')
    .trim();
}

function norm_(s) {
  return safeString_(s)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function makeCompanyRoleKey_(company, roleTitle) {
  const c = norm_(company);
  const r = norm_(roleTitle);
  if (!c && !r) return '';
  return `${c}|${r}`;
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) throw new Error('No headers found.');

  const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let c = 0; c < headerVals.length; c++) {
    const h = safeString_(headerVals[c]).trim();
    if (h) map[h] = c + 1; // 1-based
  }
  return map;
}

function ensureRequiredHeaders_(headerMap, required) {
  const missing = required.filter(h => !headerMap[h]);
  if (missing.length) throw new Error('Missing required headers: ' + missing.join(', '));
}

function setCellByHeader_(rowArray, headers, headerName, value) {
  const col = headers[headerName];
  if (!col) return;
  rowArray[col - 1] = value;
}

function safeString_(v) {
  return v == null ? '' : String(v);
}

/**
 * One-time helper: create trigger every 10 min
 */
function createJobSyncTrigger_10min() {
  ScriptApp.newTrigger('syncJobEmails_Agentic')
    .timeBased()
    .everyMinutes(10)
    .create();
}

/**
 * Test one labeled email through LLM extraction only
 */
function testOneJobEmailExtraction_() {
  const label = GmailApp.getUserLabelByName('Jobs/Inbound');
  if (!label) throw new Error('Jobs/Inbound label not found');

  const threads = label.getThreads(0, 1);
  if (!threads.length) throw new Error('No threads in Jobs/Inbound');

  const msgs = threads[0].getMessages();
  const latest = msgs[msgs.length - 1];

  const parsed = extractJobFieldsWithOpenAI_WithRetry_(
    latest.getSubject() || '',
    latest.getFrom() || '',
    (latest.getPlainBody() || '').slice(0, 5000)
  );

  Logger.log(JSON.stringify(parsed, null, 2));
}
