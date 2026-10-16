/**
 * fortheARTIST: Status Scoring (Script S1.0)
 * ---------------------------------------------------------------------------
 * Companion to the change-scoring script. This one scores an artist's STANDING
 * at a single timepoint. You score the Pre (Status) tab from pre-award material
 * and the Post (Status) tab from post-award material; the Difference tab (pure
 * formulas, no script) subtracts them and lines the result up against the
 * change rubric.
 *
 * Like the change script, this does NOT contain the rubric. It reads the wording
 * off the Rubric tab on every run, so editing the rubric never means editing code.
 *
 * ONE-TIME SETUP
 *   1. Extensions > Apps Script, paste this file, Save.
 *   2. Services (+) > Drive API > version 2 > Add.        [for PDF / Excel import]
 *   3. Reload the sheet, then FTA Status > Set or update API key.
 *   4. Select a filled-in row on Pre (Status) or Post (Status) and run
 *      FTA Status > Score active row.
 * ---------------------------------------------------------------------------
 */

// ---- settings --------------------------------------------------------------
var MODEL          = "claude-sonnet-4-6";  // or "claude-opus-4-8" for the strongest model
var MAX_TOKENS     = 4000;
var PROMPT_VERSION = "S1.0";
var API_URL        = "https://api.anthropic.com/v1/messages";

// ---- fixed layout ----------------------------------------------------------
var SHEET_RUBRIC     = "Rubric";
var SHEET_PRE        = "Pre (Status)";
var SHEET_POST       = "Post (Status)";
var SHEET_LOG        = "Log";
var N_DOMAINS        = 5;
var FIRST_DOMAIN_COL = 13;   // column M: <domain>_score
var COLS_PER_DOMAIN  = 4;    // score, evidence, rationale, evidence_sufficiency
var COL_OVERALL_NOTES_OFFSET = 2;

function isScoringSheet_(name) { return name === SHEET_PRE || name === SHEET_POST; }

// ===========================================================================
// MENU
// ===========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FTA Status")
    .addItem("Score active row", "scoreActiveRow")
    .addItem("Rescore selected rows", "rescoreSelectedRows")
    .addSeparator()
    .addItem("Import PDF into selected cell", "importPdfIntoSelectedCell")
    .addItem("Import spreadsheet into selected cell", "importSpreadsheetIntoSelectedCell")
    .addSeparator()
    .addItem("Copy scoring prompt (manual scoring)", "showScoringPrompt")
    .addItem("Paste JSON into active row", "pasteJsonIntoActiveRow")
    .addSeparator()
    .addItem("Set or update API key", "setApiKey")
    .addToUi();
}

function setApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt("Anthropic API key",
    "Paste the key (starts with sk-ant-). It is stored in this script's private properties, "
    + "not in any cell, so people viewing the sheet cannot see it.", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) { ui.alert("No key entered; nothing was saved."); return; }
  PropertiesService.getScriptProperties().setProperty("ANTHROPIC_API_KEY", key);
  ui.alert("Saved. You can now use Score active row.");
}

// ===========================================================================
// READING THE RUBRIC
// ===========================================================================
function readRubric_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RUBRIC);
  if (!sh) throw new Error('No tab named "' + SHEET_RUBRIC + '" was found in this workbook.');
  var values = sh.getDataRange().getValues();
  var version = "", headerRow = -1;
  for (var r = 0; r < values.length; r++) {
    var a = String(values[r][0] || "").trim();
    if (a.toLowerCase() === "rubric version") version = String(values[r][1] || "").trim();
    if (a.toLowerCase() === "domain" && headerRow === -1) headerRow = r;
  }
  if (headerRow === -1) throw new Error('Could not find the rubric table. Column A of the header row must say "Domain".');
  var headers = [];
  for (var c = 0; c < values[headerRow].length; c++) {
    var h = String(values[headerRow][c] || "").trim();
    if (h) headers.push(h);
  }
  var domains = [];
  for (var r2 = headerRow + 1; r2 < values.length; r2++) {
    var name = String(values[r2][0] || "").trim();
    if (!name) break;
    var row = {};
    for (var c2 = 0; c2 < headers.length; c2++) row[headers[c2]] = String(values[r2][c2] || "").trim();
    domains.push(row);
  }
  if (!domains.length) throw new Error("The rubric table has no domains in it.");
  if (!version) version = "(unversioned)";
  return { version: version, headers: headers, domains: domains };
}

/** Builds the STATUS scoring prompt from the rubric tab plus this row's evidence. */
function buildPrompt_(rubric, evidence) {
  var tp = evidence.timepoint || "this timepoint";
  var L = [];
  L.push("You are scoring ONE artist for fortheARTIST (FTA), a nonprofit that supports artists with");
  L.push("monthly guaranteed income, coaching, and pro bono professional support.");
  L.push("");
  L.push("FTA defines a sustainable career as: \"An artist has the means and agency to design a life");
  L.push("where their practice can endure and grow, supported by reliable income, practical systems for");
  L.push("managing their work and life, and the personal capacity to maintain their practice across");
  L.push("their lifetime.\"");
  L.push("");
  L.push("You are NOT judging artistic quality, talent, popularity, or output volume.");
  L.push("");
  L.push("WHAT THE SCORE MEANS");
  L.push("Every domain below is scored 1 to 5, and the scale measures the artist's STANDING at a single");
  L.push("timepoint: where the artist is right now, not how far they have moved. Score the artist's");
  L.push("current level using the exact wording of the score descriptions. Do not substitute your own scale.");
  L.push("You are scoring the " + tp.toUpperCase() + " standing for this artist.");
  L.push("");
  L.push("CALIBRATION - THE TOP OF THE SCALE IS EARNED");
  L.push("Scores of 4 and 5 are difficult to reach and describe a strong or very strong standing that is");
  L.push("well evidenced across multiple sources. Do not round upward out of goodwill or optimism. When the");
  L.push("evidence for a domain sits between two scores, choose the lower one. A 2 or 3 is a normal, honest");
  L.push("standing, not a failing one.");
  L.push("");
  L.push("ACCOUNT FOR POSITIVE RESPONSE BIAS");
  L.push("FTA gives artists unconditional monthly income and asks for nothing in return, so most artists");
  L.push("already feel warmly toward FTA and the program. Their survey and reflection language will often be");
  L.push("enthusiastic, grateful, and inflated, and that warmth is not by itself evidence of a strong standing.");
  L.push("Score the concrete, specific circumstances described, not the emotional temperature of the writing.");
  L.push("Gratitude, praise for the program, and general positivity not tied to a concrete described fact must");
  L.push("not raise a score. When the only support for a higher score is warm language, prefer the lower score");
  L.push("and say so in the rationale.");
  L.push("");
  L.push("RUBRIC (version " + rubric.version + ") - score strictly against this wording:");
  L.push("");
  for (var i = 0; i < rubric.domains.length; i++) {
    var d = rubric.domains[i];
    L.push("--------------------------------------------------------------------");
    L.push("DOMAIN " + (i + 1) + ": " + d[rubric.headers[0]]);
    for (var h = 1; h < rubric.headers.length; h++) {
      var key = rubric.headers[h];
      if (d[key]) L.push(key + ": " + d[key]);
    }
    L.push("");
  }
  L.push("--------------------------------------------------------------------");
  L.push("");
  L.push("RULES");
  L.push("- Score ONLY according to the rubric wording above. Follow it exactly.");
  L.push("- Score only the artist's standing at this timepoint, from the evidence provided. Do NOT compare");
  L.push("  to another timepoint; the pre-to-post comparison happens later on the Difference tab, not here.");
  L.push("- Base every score on the evidence text provided. NEVER invent, infer, or embellish evidence.");
  L.push("- For every domain give 1 to 3 SHORT verbatim quotes, copied exactly from the evidence text.");
  L.push("  If you cannot find a real quote for a domain, return an empty evidence list and say so in");
  L.push("  the rationale rather than paraphrasing something into quotation marks.");
  L.push("- Also give a one or two sentence rationale naming which evidence sources you relied on.");
  L.push("- The rubric folds insufficient evidence into a score of 1. So ALSO return");
  L.push('  "evidence_sufficiency": "Sufficient" or "Insufficient" for every domain, to record which');
  L.push("  situation a 1 actually was. Use \"Insufficient\" whenever the evidence is too thin to judge");
  L.push("  that domain, even if you scored above 1 on a hunch; the flag is about the evidence, not the score.");
  L.push("- Where evidence conflicts (for example warm narrative but thin specifics), say so in the rationale.");
  L.push("");
  L.push("OUTPUT - return ONLY valid JSON. No markdown, no code fences, no preamble, no commentary.");
  L.push("The domains array must contain exactly " + rubric.domains.length + " objects, in the same order as the rubric above.");
  L.push("");
  L.push("{");
  L.push('  "prompt_version": "' + PROMPT_VERSION + '",');
  L.push('  "rubric_version": "' + rubric.version + '",');
  L.push('  "artist_id": "' + (evidence.artistId || "") + '",');
  L.push('  "timepoint": "' + tp + '",');
  L.push('  "domains": [');
  for (var j = 0; j < rubric.domains.length; j++) {
    L.push('    {"domain": "' + rubric.domains[j][rubric.headers[0]] + '", "score": 0, "evidence": [""], "rationale": "", "evidence_sufficiency": "Sufficient"}'
           + (j < rubric.domains.length - 1 ? "," : ""));
  }
  L.push("  ],");
  L.push('  "overall_notes": ""');
  L.push("}");
  L.push("");
  L.push("=== EVIDENCE FOR THIS ARTIST AT THE " + tp.toUpperCase() + " TIMEPOINT ===");
  L.push("");
  L.push(evidence.text || "(none provided)");
  return L.join("\n");
}

function showScoringPrompt() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var ev = (isScoringSheet_(sheet.getName()) && row > 1) ? readEvidence_(sheet, row)
         : { artistId: "", timepoint: "", text: "" };
  var text = buildPrompt_(readRubric_(), ev);
  var html = HtmlService.createHtmlOutput(
      '<p style="font:13px Arial">Copy everything below into Claude, then paste the JSON it returns '
    + 'using <b>FTA Status &gt; Paste JSON into active row</b>.</p>'
    + '<textarea style="width:99%;height:430px;font:11px monospace">'
    + text.replace(/&/g, "&amp;").replace(/</g, "&lt;")
    + "</textarea>").setWidth(760).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, "Status scoring prompt");
}

// ===========================================================================
// SCORING
// ===========================================================================
function scoreActiveRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (!isScoringSheet_(sheet.getName())) { alert_('Switch to the "' + SHEET_PRE + '" or "' + SHEET_POST + '" tab first.'); return; }
  if (row < 2) { alert_("Select a data row (row 2 or below) first."); return; }
  scoreRow_(sheet, row, true);
}

function rescoreSelectedRows() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!isScoringSheet_(sheet.getName())) { alert_('Switch to the "' + SHEET_PRE + '" or "' + SHEET_POST + '" tab first.'); return; }
  var rng = sheet.getActiveRange();
  var first = Math.max(2, rng.getRow());
  var last = rng.getRow() + rng.getNumRows() - 1;
  var rows = [];
  for (var r = first; r <= last; r++) if (sheet.getRange(r, 1).getValue()) rows.push(r);
  if (!rows.length) { alert_("No rows with an artist_id are selected."); return; }
  var rubric = readRubric_();
  var ok = ui.alert("Rescore " + rows.length + " row(s) on " + sheet.getName() + "?",
      "They will be rescored against rubric version " + rubric.version + " and their current scores overwritten.",
      ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var done = 0, failed = [];
  for (var i = 0; i < rows.length; i++) {
    try { scoreRow_(sheet, rows[i], false); done++; }
    catch (e) { failed.push("Row " + rows[i] + ": " + e.message); }
  }
  alert_("Rescored " + done + " of " + rows.length + " row(s)." + (failed.length ? "\n\nProblems:\n" + failed.join("\n") : ""));
}

function readEvidence_(sheet, row) {
  var parts = [];
  var e1 = String(sheet.getRange(row, 10).getValue() || "");
  var e2 = String(sheet.getRange(row, 11).getValue() || "");
  var e3 = String(sheet.getRange(row, 12).getValue() || "");
  if (e1) parts.push(e1);
  if (e2) parts.push("--- COACHING NOTES ---\n" + e2);
  if (e3) parts.push("--- ADDITIONAL EVIDENCE ---\n" + e3);
  return {
    artistId:   String(sheet.getRange(row, 1).getValue() || ""),
    timepoint:  String(sheet.getRange(row, 3).getValue() || "").trim() || "this",
    sourceDocs: String(sheet.getRange(row, 4).getValue() || ""),
    text:       parts.join("\n\n")
  };
}

function scoreRow_(sheet, row, chatty) {
  var key = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!key) throw new Error("No API key saved. Run FTA Status > Set or update API key.");
  var rubric = readRubric_();
  var ev = readEvidence_(sheet, row);
  if (!ev.text) throw new Error("Row " + row + " has no evidence text. Fill the evidence columns first.");
  var prompt = buildPrompt_(rubric, ev);
  var raw = "", status = "OK";
  try {
    raw = callClaude_(key, prompt);
    var parsed = parseJson_(raw);
    validate_(parsed, rubric);
    writeScores_(sheet, row, parsed, rubric);
  } catch (err) {
    logRun_(ev, rubric, "ERROR: " + err.message, raw);
    throw err;
  }
  logRun_(ev, rubric, status, raw);
  if (chatty) alert_("Scored row " + row + " (" + ev.timepoint + ") against rubric version " + rubric.version + ".");
}

function callClaude_(key, prompt) {
  var res = UrlFetchApp.fetch(API_URL, {
    method: "post", contentType: "application/json", muteHttpExceptions: true,
    headers: { "x-api-key": key, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify({ model: MODEL, max_tokens: MAX_TOKENS, messages: [{ role: "user", content: prompt }] })
  });
  var code = res.getResponseCode(), body = res.getContentText();
  if (code === 401) throw new Error("The API key was rejected (401). Re-enter it with Set or update API key.");
  if (code === 400) throw new Error("The request was rejected (400). Check the model name in the script. Details: " + body.slice(0, 300));
  if (code === 429) throw new Error("Rate limited (429). Wait a moment and try again.");
  if (code >= 500)  throw new Error("Anthropic returned a server error (" + code + "). Try again shortly.");
  if (code !== 200) throw new Error("HTTP " + code + ": " + body.slice(0, 300));
  var data = JSON.parse(body);
  if (data.error) throw new Error(data.error.message || "Unknown API error.");
  var out = "";
  for (var i = 0; i < (data.content || []).length; i++) if (data.content[i].type === "text") out += data.content[i].text;
  if (!out) throw new Error("The model returned no text.");
  return out;
}

function parseJson_(text) {
  var t = String(text).trim().replace(/^```(?:json)?/i, "").replace(/```$/, "").trim();
  var a = t.indexOf("{"), b = t.lastIndexOf("}");
  if (a === -1 || b === -1) throw new Error("The model did not return JSON. Raw output is in the Log tab.");
  try { return JSON.parse(t.slice(a, b + 1)); }
  catch (e) { throw new Error("The JSON could not be parsed. Raw output is in the Log tab."); }
}

function validate_(p, rubric) {
  if (!p.domains || !p.domains.length) throw new Error('The response has no "domains" list.');
  if (p.domains.length !== rubric.domains.length)
    throw new Error("The rubric has " + rubric.domains.length + " domains but the model returned " + p.domains.length + ". Nothing was written.");
  if (rubric.domains.length !== N_DOMAINS)
    throw new Error("This workbook has columns for " + N_DOMAINS + " domains but the Rubric tab lists "
                    + rubric.domains.length + ". Add or remove the matching columns on the scoring tabs first.");
  for (var i = 0; i < p.domains.length; i++) {
    var s = p.domains[i].score;
    if (!(s >= 1 && s <= 5)) throw new Error("Domain " + (i + 1) + " came back with score \"" + s + "\", outside 1 to 5.");
  }
}

function writeScores_(sheet, row, p, rubric) {
  sheet.getRange(row, 5).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
  sheet.getRange(row, 6).setValue(rubric.version);
  sheet.getRange(row, 7).setValue(PROMPT_VERSION);
  sheet.getRange(row, 8).setValue(MODEL);
  sheet.getRange(row, 9).setValue(Session.getActiveUser().getEmail() || "Claude (API)");
  for (var i = 0; i < p.domains.length; i++) {
    var d = p.domains[i], c = FIRST_DOMAIN_COL + i * COLS_PER_DOMAIN;
    var quotes = Array.isArray(d.evidence) ? d.evidence.filter(String).join("  |  ") : String(d.evidence || "");
    sheet.getRange(row, c).setValue(d.score);
    sheet.getRange(row, c + 1).setValue(quotes);
    sheet.getRange(row, c + 2).setValue(d.rationale || "");
    sheet.getRange(row, c + 3).setValue(d.evidence_sufficiency === "Insufficient" ? "Insufficient" : "Sufficient");
  }
  var compositeCol = FIRST_DOMAIN_COL + N_DOMAINS * COLS_PER_DOMAIN;
  sheet.getRange(row, compositeCol + COL_OVERALL_NOTES_OFFSET).setValue(p.overall_notes || "");
  SpreadsheetApp.flush();
}

function logRun_(ev, rubric, status, raw) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (!sh) return;
  sh.appendRow([new Date(), ev.artistId, ev.timepoint, ev.sourceDocs, MODEL, PROMPT_VERSION,
                rubric ? rubric.version : "", Session.getActiveUser().getEmail() || "(unknown)",
                status, String(raw).slice(0, 45000)]);
}

// ===========================================================================
// MANUAL PATH
// ===========================================================================
function pasteJsonIntoActiveRow() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  if (!isScoringSheet_(sheet.getName()) || row < 2) { alert_('Select a data row on Pre (Status) or Post (Status) first.'); return; }
  var res = ui.prompt("Paste JSON", "Paste the whole JSON block Claude returned.", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var rubric = readRubric_(), ev = readEvidence_(sheet, row);
  try {
    var parsed = parseJson_(res.getResponseText());
    validate_(parsed, rubric);
    writeScores_(sheet, row, parsed, rubric);
    sheet.getRange(row, 8).setValue("manual paste");
  } catch (err) {
    logRun_(ev, rubric, "ERROR (manual paste): " + err.message, res.getResponseText());
    alert_(err.message); return;
  }
  logRun_(ev, rubric, "OK (manual paste)", res.getResponseText());
  alert_("Row " + row + " filled in from the pasted JSON.");
}

// ===========================================================================
// DRIVE IMPORT
// ===========================================================================
function driveIdFromLinkOrId_(value) {
  var m = String(value).trim().match(/[-\w]{25,}/);
  return m ? m[0] : String(value).trim();
}

function convertPdfToText_(fileId) {
  var copied = Drive.Files.copy({ title: "__fta_tmp_pdf", mimeType: "application/vnd.google-apps.document" }, fileId, { supportsAllDrives: true });
  try { return DocumentApp.openById(copied.id).getBody().getText(); }
  finally { DriveApp.getFileById(copied.id).setTrashed(true); }
}

function spreadsheetToText_(fileId) {
  var copied = Drive.Files.copy({ title: "__fta_tmp_sheet", mimeType: "application/vnd.google-apps.spreadsheet" }, fileId, { supportsAllDrives: true });
  try {
    var data = SpreadsheetApp.openById(copied.id).getSheets()[0].getDataRange().getValues();
    if (!data.length) return "";
    var headers = data[0].map(String), typeCol = -1;
    for (var h = 0; h < headers.length; h++) if (headers[h].toLowerCase().indexOf("report type") >= 0) typeCol = h;
    var blocks = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r], any = false, lines = [];
      for (var c = 0; c < headers.length; c++) {
        var v = row[c];
        if (v !== "" && v !== null) { lines.push(headers[c] + ": " + String(v)); any = true; }
      }
      if (!any) continue;
      var label = (typeCol >= 0 && row[typeCol]) ? "=== " + String(row[typeCol]) + " ===" : "=== Record " + r + " ===";
      blocks.push(label + "\n" + lines.join("\n"));
    }
    return blocks.join("\n\n");
  } finally { DriveApp.getFileById(copied.id).setTrashed(true); }
}

function importPdfIntoSelectedCell()         { importIntoSelectedCell_("pdf"); }
function importSpreadsheetIntoSelectedCell() { importIntoSelectedCell_("sheet"); }

function importIntoSelectedCell_(kind) {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  if (!isScoringSheet_(sheet.getName()) || cell.getRow() < 2) {
    alert_("Click the evidence cell you want the text to land in first, on Pre (Status) or Post (Status).");
    return;
  }
  var res = ui.prompt(kind === "pdf" ? "Import a PDF" : "Import a spreadsheet",
    "Paste the Google Drive link (or the file ID).\n\nThe text will go into cell " + cell.getA1Notation() + ".",
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var link = res.getResponseText().trim();
  if (!link) return;
  var id = driveIdFromLinkOrId_(link), text, name;
  try {
    name = DriveApp.getFileById(id).getName();
    text = (kind === "pdf") ? convertPdfToText_(id) : spreadsheetToText_(id);
  } catch (e) {
    alert_("Could not read that file.\n\n" + e.message + "\n\nCheck the link points to a file you can open, and that Drive API v2 is added under Services.");
    return;
  }
  if (!String(text).trim()) { alert_("That file opened but produced no text. If it is a scanned PDF, the scan may be too poor to read."); return; }
  cell.setValue(text);
  var docCell = sheet.getRange(cell.getRow(), 4);
  var existing = String(docCell.getValue() || "");
  if (existing.indexOf(name) === -1) docCell.setValue(existing ? existing + "; " + name : name);
  alert_("Imported " + String(text).length + " characters from \"" + name + "\" into " + cell.getA1Notation() + ".\n\nRead it over before scoring.");
}

// ---------------------------------------------------------------------------
function alert_(msg) { SpreadsheetApp.getUi().alert(msg); }