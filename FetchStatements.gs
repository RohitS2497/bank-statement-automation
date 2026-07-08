/**
 * FetchStatements.gs  —  Gmail ➜ Drive automation (runs inside YOUR Google account)
 * ---------------------------------------------------------------------------------
 * Stage 1 of the pipeline: FETCH + STORE.
 * Searches Gmail for bank-statement emails, saves the PDF attachments into an
 * organised Drive folder tree, prevents duplicates, and logs every file to a Sheet.
 *
 * Setup (5 minutes):
 *   1. Go to script.google.com  ➜  New project  ➜  paste this file.
 *   2. Edit the CONFIG block below (Gmail search + optional root folder name).
 *   3. Run `fetchBankStatements` once and click "Allow" to grant Gmail+Drive access.
 *   4. Run `installTrigger` once to make it run automatically every hour.
 */

// ------------------------------- CONFIG -------------------------------
const CONFIG = {
  // Gmail search. Tighten this to your senders/label for precision.
  //   Examples: 'has:attachment filename:pdf subject:statement'
  //             'has:attachment filename:pdf from:(alerts@hdfcbank.net OR noreply@icicibank.com)'
  //             'label:bank-statements has:attachment filename:pdf'
  SEARCH_QUERY: 'has:attachment filename:pdf (subject:statement OR subject:"account statement")',
  ROOT_FOLDER: 'Bank Statements',   // top-level Drive folder that gets created
  PROCESSED_LABEL: 'statement-saved', // Gmail label added to processed threads
  MAX_THREADS: 50,                  // safety cap per run
  ORGANISE_BY: 'year-month'         // 'year' | 'year-month' | 'flat'
};

// ------------------------------- MAIN -------------------------------
function fetchBankStatements() {
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), CONFIG.ROOT_FOLDER);
  const label = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  const sheet = getOrCreateLogSheet_(root);

  const threads = GmailApp.search(CONFIG.SEARCH_QUERY, 0, CONFIG.MAX_THREADS);
  Logger.log('Found %s thread(s) matching query.', threads.length);

  let saved = 0, skipped = 0;
  threads.forEach(function (thread) {
    // Skip threads we've already processed (idempotent / no duplicates)
    if (thread.getLabels().some(l => l.getName() === CONFIG.PROCESSED_LABEL)) return;

    thread.getMessages().forEach(function (msg) {
      const when = msg.getDate();
      msg.getAttachments().forEach(function (att) {
        if (att.getContentType() !== 'application/pdf' &&
            !/\.pdf$/i.test(att.getName())) return;

        const folder = targetFolder_(root, when);
        const fileName = systematicName_(when, msg.getFrom(), att.getName());

        // Duplicate prevention: skip if a file with this name already exists
        if (folder.getFilesByName(fileName).hasNext()) { skipped++; return; }

        const file = folder.createFile(att.copyBlob()).setName(fileName);
        sheet.appendRow([new Date(), fileName, msg.getFrom(), when,
                         folder.getName(), file.getUrl()]);
        saved++;
      });
    });
    thread.addLabel(label);   // mark processed
    thread.markRead();
  });

  Logger.log('Done. Saved %s new file(s), skipped %s duplicate(s).', saved, skipped);
  return { saved: saved, skipped: skipped };
}

// ------------------------------- HELPERS -------------------------------
function targetFolder_(root, when) {
  if (CONFIG.ORGANISE_BY === 'flat') return root;
  const y = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy');
  const yFolder = getOrCreateFolder_(root, y);
  if (CONFIG.ORGANISE_BY === 'year') return yFolder;
  const ym = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM');
  return getOrCreateFolder_(yFolder, ym);
}

function systematicName_(when, from, original) {
  const d = Utilities.formatDate(when, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sender = (from.match(/@([\w.-]+)/) || [,'unknown'])[1].split('.')[0];
  const clean = original.replace(/[^\w.\- ]+/g, '_').replace(/\s+/g, '_');
  return d + '__' + sender + '__' + clean;      // e.g. 2026-05-31__icicibank__May_Statement.pdf
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function getOrCreateLogSheet_(root) {
  const name = '_StatementLog';
  const it = root.getFilesByName(name);
  let ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next());
  } else {
    ss = SpreadsheetApp.create(name);
    DriveApp.getFileById(ss.getId()).moveTo(root);
    ss.getActiveSheet().appendRow(
      ['Logged At', 'Saved File', 'From', 'Email Date', 'Folder', 'Drive URL']);
  }
  return ss.getActiveSheet();
}

// ------------------------------- TRIGGER -------------------------------
function installTrigger() {
  // remove existing triggers for this function, then add one hourly trigger
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'fetchBankStatements') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('fetchBankStatements').timeBased().everyHours(1).create();
  Logger.log('Hourly trigger installed.');
}
