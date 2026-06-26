/*
 * Fountainize test suite — runs in the REAL Apps Script environment.
 *
 * It creates a real Google Doc, runs the real convert() on Fountain-format input,
 * reads back the ACTUAL rendered formatting (indent, bold, font, heading,
 * alignment, text) and asserts it. No mocks / proxies.
 *
 * HOW TO RUN
 *   1. Open the script editor, select `runFountainTests`, click Run.
 *   2. Authorize the (documents) scope when prompted.
 *   3. Read the PASS/FAIL report in the Execution log, and open the "Fountainize
 *      Test Sample" doc it creates to eyeball the rendered result.
 *
 * Requires the broader `https://www.googleapis.com/auth/documents` scope (to
 * create/read a scratch doc) — this is a dev/test tool, not used by the add-on.
 */

var PT = 72; // points per inch

// ---- expected indents (inches -> points), per Style in Fountainize.js ----
var IND = { scene: 0, action: 0, transition: 0, character: 2 * PT, dialogue: 1 * PT, parenthetical: 1.5 * PT };

// Read the actual rendered properties of a paragraph and infer its element type.
function readEl(p) {
  var ind = p.getIndentStart(); ind = (ind == null) ? 0 : ind;
  var sa = p.getSpacingAfter(); sa = (sa == null) ? 0 : sa;
  var sb = p.getSpacingBefore(); sb = (sb == null) ? 0 : sb;
  var align = p.getAlignment();
  var isH3 = p.getHeading() === DocumentApp.ParagraphHeading.HEADING3;
  var bold = null, font = null;
  try { bold = p.editAsText().isBold(); font = p.editAsText().getFontFamily(); } catch (e) {}
  var near = function (a, b) { return Math.abs(a - b) <= 3; };
  var type;
  if (isH3) type = 'scene';
  else if (align === DocumentApp.HorizontalAlignment.RIGHT) type = 'transition';
  else if (align === DocumentApp.HorizontalAlignment.CENTER) type = 'centered';
  else if (near(ind, IND.character)) type = 'character';
  else if (near(ind, IND.parenthetical)) type = 'parenthetical';
  else if (near(ind, IND.dialogue)) type = 'dialogue';
  else type = 'action';
  return { text: p.getText(), type: type, indent: ind, spaceBefore: sb, spaceAfter: sa, bold: bold, font: font, isH3: isH3, align: align };
}

function isBlankPara(p) { var t = p.getText(); return !t || t.trim() === ''; }

// Reset the scratch doc to a single clean (Normal) paragraph, then write the lines.
function buildDoc(body, lines) {
  body.clear();
  var ps = body.getParagraphs();
  var first = ps.length ? ps[0] : body.appendParagraph('');
  first.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  first.setIndentStart(0); first.setIndentFirstLine(0); first.setIndentEnd(0);
  first.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  first.editAsText().setBold(false);
  if (!lines.length) { first.setText(''); return; }
  first.setText(lines[0]);
  for (var i = 1; i < lines.length; i++) body.appendParagraph(lines[i]);
}

// Run convert() on `lines` in the scratch doc; return {chars, els:[readEl for non-blank], all:[readEl for every para]}.
function runOn(doc, lines, opts) {
  opts = opts || {};
  var body = doc.getBody();
  buildDoc(body, lines);
  activeDocOverride = doc;
  var chars;
  try {
    chars = convert('whole', !!opts.sceneNumbers, opts.autoFonts !== false, true);
  } finally {
    activeDocOverride = null;
  }
  var paras = body.getParagraphs();
  var els = [], all = [];
  for (var i = 0; i < paras.length; i++) {
    var r = readEl(paras[i]);
    all.push(r);
    if (!isBlankPara(paras[i])) els.push(r);
  }
  return { chars: chars, els: els, all: all };
}

function typesOf(els) { return els.map(function (e) { return e.type; }); }
function arrEq(a, b) { return a.length === b.length && a.every(function (x, i) { return x === b[i]; }); }

// ---------- implemented-feature tests ----------
// Each: name, input lines, expected element types (non-blank), optional extra(els, res)->''|errorString
var TESTS = [
  { name: 'Scene heading INT.', input: ['INT. KITCHEN - DAY', 'Bob pours coffee.'], expect: ['scene', 'action'],
    extra: function (e) { return (e[0].bold === true && e[0].isH3 && e[0].font === 'Courier Prime' && e[0].text === 'INT. KITCHEN - DAY') ? '' : 'scene not bold/H3/CourierPrime/uppercase: ' + JSON.stringify(e[0]); } },
  { name: 'Scene heading EXT.', input: ['EXT. PARK - NIGHT', 'Leaves rustle.'], expect: ['scene', 'action'] },
  { name: 'Combined scene heading EXT/INT. (prefix preserved)', input: ['EXT/INT. CAR - DAY', 'She drives fast.'], expect: ['scene', 'action'],
    extra: function (e) { return e[0].text === 'EXT/INT. CAR - DAY' ? '' : 'prefix not preserved: ' + e[0].text; } },
  { name: 'Action', input: ['The storm rolls in over the hills.'], expect: ['action'] },
  { name: 'Character + Dialogue', input: ['JANE', 'I told you not to come.'], expect: ['character', 'dialogue'] },
  { name: 'Parenthetical', input: ['JANE', '(whispering)', 'They can hear us.'], expect: ['character', 'parenthetical', 'dialogue'] },
  { name: 'Transition (right-aligned)', input: ['He slams the door.', 'CUT TO:', 'INT. HALL - DAY'], expect: ['action', 'transition', 'scene'],
    extra: function (e) { return e[1].align === DocumentApp.HorizontalAlignment.RIGHT ? '' : 'transition not right-aligned'; } },
  { name: 'Centered text >..< (brackets stripped, centered)', input: ['>THE END<'], expect: ['centered'],
    extra: function (e) { return (e[0].text === 'THE END' && e[0].align === DocumentApp.HorizontalAlignment.CENTER) ? '' : 'centered wrong: ' + JSON.stringify(e[0]); } },
  { name: 'Character with extension (V.O.)', input: ['JANE (V.O.)', 'It was a dark night.'], expect: ['character', 'dialogue'] },
  { name: 'Character cue with number/symbol', input: ['GUARD #1', 'Halt!'], expect: ['character', 'dialogue'] },
  { name: 'Inline parenthetical dialogue does not swallow next action', input: ['JANE', '(beat) Fine.', 'She storms off.'], expect: ['character', 'dialogue', 'action'] },
  { name: 'All-caps action line is not mistaken for a transition/cue context', input: ['EXT. ROOF - DAY', 'He looks down.', 'CUT TO:', 'EXT. STREET - DAY'], expect: ['scene', 'action', 'transition', 'scene'] },
  { name: 'Character auto-detection populates the shortcut list', input: ['SAMANTHA', 'Hello.'], expect: ['character', 'dialogue'],
    extra: function (e, res) { var names = (res.chars || []).map(function (c) { return c.name; }); return names.indexOf('SAMANTHA') > -1 ? '' : 'SAMANTHA not auto-detected: ' + JSON.stringify(res.chars); } },
  { name: 'Scene numbers option prefixes the heading', input: ['INT. ROOM - DAY', 'Wait.'], opts: { sceneNumbers: true }, expect: ['scene', 'action'],
    extra: function (e) { return (e[0].text.indexOf('1') === 0 && e[0].text.indexOf('INT. ROOM - DAY') > -1) ? '' : 'no scene number prefix: ' + e[0].text; } },
];

// ---------- spacing + idempotency + selection (checked separately) ----------
// Spacing is paragraph "space before"/"space after" MARGINS (points), not blank lines.
// Per type: scene/action/transition/centered 12 before+after; character 12 before, 6 after;
// dialogue 0 before, 6 after; parenthetical 0/0. And zero blank paragraphs.
function checkSpacing(doc) {
  var res = runOn(doc, ['INT. ROOM - DAY', 'He waits.', 'JANE', 'Sit.', 'She sits down.']); // scene, action, character, dialogue, action
  var blanks = res.all.filter(function (p) { return !p.text || p.text.trim() === ''; }).length;
  var before = res.els.map(function (e) { return e.spaceBefore; });
  var after = res.els.map(function (e) { return e.spaceAfter; });
  var wantBefore = [BASE_GAP * 2, BASE_GAP, BASE_GAP, 0, BASE_GAP]; // scene gets a double gap above
  var wantAfter = [BASE_GAP, BASE_GAP, BASE_GAP / 2, BASE_GAP / 2, BASE_GAP];
  return { ok: blanks === 0 && arrEq(before, wantBefore) && arrEq(after, wantAfter),
           blanks: blanks, before: before, after: after, wantBefore: wantBefore, wantAfter: wantAfter };
}

// Full fingerprint: one entry per paragraph with type, space before/after and text. A
// second Format Script pass must not change this (no extra lines/spaces/headers).
function structure(doc) {
  return doc.getBody().getParagraphs().map(function (p) {
    if (isBlankPara(p)) { return 'blank'; }
    var r = readEl(p);
    return r.type + '|sb=' + r.spaceBefore + '|sa=' + r.spaceAfter + '|' + p.getText();
  }).join('\n');
}

function checkIdempotent(doc) {
  var lines = ['INT. ROOM - DAY', 'He waits.', 'JANE', 'Sit.', 'She sits down.', 'CUT TO:', 'EXT. STREET - DAY', 'A car passes by.'];
  runOn(doc, lines);
  var s1 = structure(doc);
  activeDocOverride = doc;
  try { convert('whole', false, true, true); } finally { activeDocOverride = null; } // re-run must be a no-op
  var s2 = structure(doc);
  return { ok: s1 === s2, first: s1, second: s2 };
}

// formatSelection() formats only the selected paragraphs and leaves the rest alone.
function checkSelection(doc) {
  var body = doc.getBody();
  buildDoc(body, ['Untouched action above.', 'JANE', 'Hello there.', 'Untouched line below.']);
  var paras = body.getParagraphs();
  doc.setSelection(doc.newRange().addElement(paras[1]).addElement(paras[2]).build()); // select JANE + dialogue
  activeDocOverride = doc;
  try { formatSelection(false, false, true); } finally { activeDocOverride = null; }

  var jane, dlg, below;
  body.getParagraphs().forEach(function (p) {
    var t = p.getText();
    if (t === 'JANE') jane = readEl(p);
    else if (t.indexOf('Hello there') === 0) dlg = readEl(p);
    else if (t.indexOf('Untouched line below') === 0) below = readEl(p);
  });
  var ok = jane && jane.type === 'character' && dlg && dlg.type === 'dialogue'
        && below && below.spaceAfter === 0; // out-of-selection paragraph left untouched (no margin applied)
  return { ok: !!ok, jane: jane && jane.type, dlg: dlg && dlg.type, belowSpaceAfter: below && below.spaceAfter };
}

// Formatting must not strip a user's manual bold from non-scene lines (e.g. episode titles).
function checkBoldPreserved(doc) {
  var body = doc.getBody();
  buildDoc(body, ['INT. ROOM - DAY', 'This action line is bold.']);
  body.getParagraphs()[1].editAsText().setBold(true); // user-bolded action line
  activeDocOverride = doc;
  try { convert('whole', false, false, true); } finally { activeDocOverride = null; }
  var line = body.getParagraphs()[1];
  var bold = line.editAsText().isBold();
  return { ok: bold === true, bold: bold };
}

// A user-set heading (episode title, H2) is kept as H2, made bold, with a 24pt top gap.
function checkEpisodeTitle(doc) {
  var body = doc.getBody();
  buildDoc(body, ['Episode One', 'INT. ROOM - DAY', 'She enters.']);
  body.getParagraphs()[0].setHeading(DocumentApp.ParagraphHeading.HEADING2); // user marks it H2
  activeDocOverride = doc;
  try { convert('whole', false, false, true); } finally { activeDocOverride = null; }
  var t = body.getParagraphs()[0];
  var bold = t.editAsText().isBold();
  var sb = t.getSpacingBefore();
  var stillH2 = t.getHeading() === DocumentApp.ParagraphHeading.HEADING2;
  return { ok: bold === true && sb === BASE_GAP * 2 && stillH2, bold: bold, spaceBefore: sb, h2: stillH2 };
}

// Boneyard /* ... */ blocks are coloured blue but keep their position-based formatting,
// while text outside stays black. And re-running keeps the colours stable (idempotent).
function checkBoneyard(doc) {
  var body = doc.getBody();
  buildDoc(body, ['Action before.', '/*', 'INT. CUT SCENE - DAY', 'Commented action.', '*/', 'Action after.']);
  var BLUE = '#1155cc', BLACK = '#000000';
  var colorOf = function (prefix) {
    var ps = body.getParagraphs();
    for (var i = 0; i < ps.length; i++) {
      if (ps[i].getText().indexOf(prefix) === 0) { try { return ps[i].editAsText().getForegroundColor(); } catch (e) { return null; } }
    }
    return null;
  };
  var typeOf = function (prefix) {
    var ps = body.getParagraphs();
    for (var i = 0; i < ps.length; i++) { if (ps[i].getText().indexOf(prefix) === 0) { return readEl(ps[i]).type; } }
    return null;
  };
  var run = function () { activeDocOverride = doc; try { convert('whole', false, false, true); } finally { activeDocOverride = null; } };
  run();
  var blueInside = colorOf('/*') === BLUE && colorOf('INT. CUT SCENE') === BLUE && colorOf('Commented') === BLUE && colorOf('*/') === BLUE;
  var blackOutside = colorOf('Action before') === BLACK && colorOf('Action after') === BLACK;
  var sceneKept = typeOf('INT. CUT SCENE') === 'scene'; // formatting unchanged by position
  run(); // re-run must keep colours stable
  var stable = colorOf('INT. CUT SCENE') === BLUE && colorOf('Action after') === BLACK;
  return { ok: blueInside && blackOutside && sceneKept && stable, blueInside: blueInside, blackOutside: blackOutside, sceneKept: sceneKept, stable: stable };
}

// ---------- Fountain features NOT implemented (flagged, not failed) ----------
var UNIMPLEMENTED = [
  { feature: 'Forced Scene Heading', markup: '.SNOWGLOBE', fountain: 'leading "." forces a scene heading' },
  { feature: 'Forced Action', markup: '!CAPS THAT WOULD READ AS A CUE', fountain: 'leading "!" forces action' },
  { feature: 'Forced Character', markup: '@McCLANE', fountain: 'leading "@" forces a character cue (even lowercase)' },
  { feature: 'Forced Transition', markup: '> JARRING SMASH CUT', fountain: 'leading ">" (no closing "<") forces a transition' },
  { feature: 'Dual Dialogue', markup: 'BRICK ^', fountain: 'trailing "^" makes side-by-side dialogue' },
  { feature: 'Lyrics', markup: '~Willy Wonka! Willy Wonka!', fountain: 'leading "~" marks lyrics' },
  { feature: 'Emphasis (bold/italic/underline)', markup: 'He was *very* **angry** and _tired_', fountain: '*italic* **bold** _underline_ ***bolditalic***' },
  { feature: 'Page Break', markup: '===', fountain: 'three or more "=" forces a page break' },
  { feature: 'Note', markup: '[[ remember the red herring ]]', fountain: '"[[ .. ]]" is an inline note (hidden in output)' },
  { feature: 'Boneyard (comment)', markup: '/* this whole bit is cut */', fountain: '"/* .. */" is normally removed from the script; Fountainize instead keeps it and colours the block blue (see the Boneyard test)' },
  { feature: 'Section', markup: '# Act One', fountain: 'leading "#" is a structural section (outline only)' },
  { feature: 'Synopsis', markup: '= Sara reaches the gate', fountain: 'leading "=" is a synopsis (outline only)' },
  { feature: 'Title Page (markup)', markup: 'Title: Big Fish', fountain: '"Title:/Author:/.." key:value title page (Fountainize uses the sidebar form instead)' }
];

function runUnimplemented(doc) {
  return UNIMPLEMENTED.map(function (u) {
    var got = '(error)';
    try {
      var res = runOn(doc, [u.markup, 'A following line of normal text.']);
      got = res.els.length ? res.els[0].type : '(none)';
    } catch (err) { got = 'EXCEPTION: ' + err; }
    return { feature: u.feature, markup: u.markup, fountain: u.fountain, got: got };
  });
}

// ---------- runner ----------
// Web-app entry point: triggers the suite via a plain HTTP GET (returns the text
// report) once the script has been authorized. Used for automated test runs.
function doGet(e) {
  var out;
  try { out = runFountainTests(); }
  catch (err) { return ContentService.createTextOutput('RUN ERROR: ' + err).setMimeType(ContentService.MimeType.TEXT); }
  return ContentService.createTextOutput(out.report + '\n\nsampleDoc: ' + out.sampleDocUrl).setMimeType(ContentService.MimeType.TEXT);
}

function runFountainTests() {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var doc = DocumentApp.create('Fountainize Test Sample ' + stamp);
  var lines = [];
  var pass = 0, fail = 0;

  lines.push('FOUNTAINIZE TEST REPORT  —  ' + stamp);
  lines.push('Verifies the rendered output of the real convert() on a real Google Doc.');
  lines.push('');
  lines.push('== IMPLEMENTED FEATURES ==');
  TESTS.forEach(function (t) {
    var got = [], ok = false, detail = '';
    try {
      var res = runOn(doc, t.input, t.opts);
      got = typesOf(res.els);
      ok = arrEq(got, t.expect);
      if (ok && t.extra) { detail = t.extra(res.els, res); if (detail) ok = false; }
    } catch (err) { ok = false; detail = 'EXCEPTION: ' + err; }
    if (ok) pass++; else fail++;
    lines.push((ok ? 'PASS  ' : 'FAIL  ') + t.name);
    if (!ok) {
      lines.push('        input:    ' + JSON.stringify(t.input));
      lines.push('        expected: ' + JSON.stringify(t.expect));
      lines.push('        got:      ' + JSON.stringify(got));
      if (detail) lines.push('        check:    ' + detail);
    }
  });

  try {
    var sp = checkSpacing(doc);
    (sp.ok ? pass++ : fail++);
    lines.push((sp.ok ? 'PASS  ' : 'FAIL  ') + 'Spacing: before/after margins, no blank lines  (before ' + JSON.stringify(sp.before) + ' want ' + JSON.stringify(sp.wantBefore) + '; after ' + JSON.stringify(sp.after) + ' want ' + JSON.stringify(sp.wantAfter) + '; blanks ' + sp.blanks + ')');
  } catch (err) { fail++; lines.push('FAIL  Spacing — EXCEPTION: ' + err); }

  try {
    var idem = checkIdempotent(doc);
    (idem.ok ? pass++ : fail++);
    lines.push((idem.ok ? 'PASS  ' : 'FAIL  ') + 'Idempotent: a second Format Script pass changes nothing');
    if (!idem.ok) { lines.push('        first:  ' + JSON.stringify(idem.first)); lines.push('        second: ' + JSON.stringify(idem.second)); }
  } catch (err) { fail++; lines.push('FAIL  Idempotent — EXCEPTION: ' + err); }

  try {
    var sel = checkSelection(doc);
    (sel.ok ? pass++ : fail++);
    lines.push((sel.ok ? 'PASS  ' : 'FAIL  ') + 'Format selection: styles only the selection  (jane=' + sel.jane + ', dlg=' + sel.dlg + ', belowSpaceAfter=' + sel.belowSpaceAfter + ')');
  } catch (err) { fail++; lines.push('FAIL  Format selection — EXCEPTION: ' + err); }

  try {
    var bp = checkBoldPreserved(doc);
    (bp.ok ? pass++ : fail++);
    lines.push((bp.ok ? 'PASS  ' : 'FAIL  ') + 'Bold preserved: manual bold on non-scene lines is kept  (bold=' + bp.bold + ')');
  } catch (err) { fail++; lines.push('FAIL  Bold preserved — EXCEPTION: ' + err); }

  try {
    var et = checkEpisodeTitle(doc);
    (et.ok ? pass++ : fail++);
    lines.push((et.ok ? 'PASS  ' : 'FAIL  ') + 'Episode title (H2): kept as H2, bold, 24pt top gap  (bold=' + et.bold + ', spaceBefore=' + et.spaceBefore + ', h2=' + et.h2 + ')');
  } catch (err) { fail++; lines.push('FAIL  Episode title — EXCEPTION: ' + err); }

  try {
    var by = checkBoneyard(doc);
    (by.ok ? pass++ : fail++);
    lines.push((by.ok ? 'PASS  ' : 'FAIL  ') + 'Boneyard /* */: block coloured blue, formatting kept, idempotent  (blueInside=' + by.blueInside + ', blackOutside=' + by.blackOutside + ', sceneKept=' + by.sceneKept + ', stable=' + by.stable + ')');
  } catch (err) { fail++; lines.push('FAIL  Boneyard — EXCEPTION: ' + err); }

  lines.push('');
  lines.push('== NOT IMPLEMENTED (Fountain spec) — flagged, showing what Fountainize currently does ==');
  runUnimplemented(doc).forEach(function (u) {
    lines.push('TODO  ' + u.feature + '  (' + u.markup + ')');
    lines.push('        fountain: ' + u.fountain);
    lines.push('        current:  Fountainize renders this line as "' + u.got + '"');
  });

  lines.push('');
  lines.push('== SUMMARY ==  ' + pass + ' passed, ' + fail + ' failed, ' + UNIMPLEMENTED.length + ' Fountain features not implemented.');

  var report = lines.join('\n');
  Logger.log(report); // always available in the editor's Execution log

  // Leave a rendered mini-screenplay in the doc so you can visually confirm it reads
  // like a screenplay on the actual page.
  try {
    runOn(doc, [
      'INT. NOMAD CAMP - DAWN', 'Sara wakes in a small tent and peeks outside.',
      'ZE\'EV', '(quietly)', 'The road to the city is open.',
      'She nods and reaches for her rifle.', 'CUT TO:', 'EXT. DESERT ROAD - DAY',
      'The tribe walks along a ruined highway.', 'DOV', 'Something is not right.', '>THE END<'
    ]);
  } catch (e) {}

  return { summary: pass + ' passed / ' + fail + ' failed', sampleDocUrl: doc.getUrl(), report: report };
}
