// ============================================================================
// Fountainize core — classify each paragraph and apply screenplay formatting.
//
// Speed: each element is formatted with ONE batched el.setAttributes() call
// (indents + spacing + line height + bold + font + colour + alignment) instead
// of many. Gaps are paragraph "space after" MARGINS, not blank paragraphs, so
// there is no insert/remove churn and a second Format Script pass is a no-op.
// ============================================================================

var doc, body, pArray;            // active document / its body / paragraph snapshot
var charList = [];                // [{name, sh}] character shorthands
var version = 10;

var PT = 72;                      // points per inch
var BASE_GAP = 12;                // points of "space after" a block — the one spacing knob
var LINE_HEIGHT = 0.86;           // line spacing
var screenplayFont = 'Courier Prime';

var activeDocOverride = null;     // tests can point convert()/docSetUp() at a document
function getActiveDoc(){
  return activeDocOverride ? activeDocOverride : DocumentApp.getActiveDocument();
}

// ---- Element styles --------------------------------------------------------
// A Style precomputes its full DocumentApp.Attribute map so apply() is one write.
var LEFT   = DocumentApp.HorizontalAlignment.LEFT;
var RIGHT  = DocumentApp.HorizontalAlignment.RIGHT;
var CENTER = DocumentApp.HorizontalAlignment.CENTER;
var HEADING3 = DocumentApp.ParagraphHeading.HEADING3;

function Style(o){
  this.uCase   = !!o.uCase;
  this.heading = o.heading || null;            // scenes only (Heading 3 → document outline)
  var a = {};
  a[DocumentApp.Attribute.INDENT_START]         = o.iLeft * PT;
  a[DocumentApp.Attribute.INDENT_FIRST_LINE]    = o.iLeft * PT;
  a[DocumentApp.Attribute.INDENT_END]           = o.iRight * PT;
  a[DocumentApp.Attribute.LINE_SPACING]         = LINE_HEIGHT;
  a[DocumentApp.Attribute.SPACING_BEFORE]       = 0;
  a[DocumentApp.Attribute.SPACING_AFTER]        = o.spaceAfter;
  a[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = o.align || LEFT;
  a[DocumentApp.Attribute.BOLD]                 = !!o.bold;
  a[DocumentApp.Attribute.FONT_FAMILY]          = screenplayFont;
  a[DocumentApp.Attribute.FONT_SIZE]            = 12;
  a[DocumentApp.Attribute.FOREGROUND_COLOR]     = '#000000';
  this.attrs = a;
}

// indents in inches; spaceAfter in points (https://screenwriting.io/what-is-standard-screenplay-format/)
var scene            = new Style({iLeft:0,    iRight:0,   uCase:true,  spaceAfter:BASE_GAP,   bold:true, heading:HEADING3});
var sceneWithNumbers = new Style({iLeft:-0.5, iRight:0,   uCase:true,  spaceAfter:BASE_GAP,   bold:true, heading:HEADING3});
var action           = new Style({iLeft:0,    iRight:0,   uCase:false, spaceAfter:BASE_GAP});
var character        = new Style({iLeft:2.0,  iRight:0,   uCase:true,  spaceAfter:BASE_GAP});
var dialogue         = new Style({iLeft:1.0,  iRight:1.5, uCase:false, spaceAfter:BASE_GAP/2});
var parenthetical    = new Style({iLeft:1.5,  iRight:1.9, uCase:false, spaceAfter:BASE_GAP/2});
var transition       = new Style({iLeft:0,    iRight:0,   uCase:true,  spaceAfter:BASE_GAP,   align:RIGHT});
var centered         = new Style({iLeft:-0.5, iRight:0,   uCase:false, spaceAfter:BASE_GAP,   align:CENTER});

// Apply a style to one paragraph: uppercase (only if it changes), set the scene
// heading (scenes), then ONE batched setAttributes carrying everything else.
function apply(el, style){
  if(style.uCase){
    var t = el.getText(), up = t.toUpperCase();
    if(up !== t){ el.setText(up); }
  }
  if(style.heading){ el.setHeading(style.heading); } // before setAttributes so our attrs win over the named style
  el.setAttributes(style.attrs);
}

// ---- Main: classify every paragraph, then apply ----------------------------
// type: 'whole' (default) formats the body; 'selection' formats the highlighted
// paragraphs (or the one under the cursor). Returns the character shorthand list.
function convert(type, sceneNumbers, autoFontsMargins, endPunctuationMeansNotChar){
  charList = getCharsFromStorage();
  var charLimit = isLicenseValid() ? 15 : 5;
  var sceneNum = 1;

  doc = getActiveDoc();
  body = doc.getBody();
  type = type || 'whole';

  if(autoFontsMargins){ docSetUp(); }

  pArray = collectParagraphs(type);
  if(pArray === null){ return; } // selection requested but nothing to format (already alerted)

  // Free-tier length cap (whole-doc only).
  if(type === 'whole' && !isLicenseValid() && pArray.length > 800){
    DocumentApp.getUi().alert("Exceeded Element Limit",
      "Looks like your script is getting pretty long (" + pArray.length + " elements). Fountainize Free limits your script to 800 elements. Upgrade to Pro for unlimited length!",
      DocumentApp.getUi().ButtonSet.OK);
    return false;
  }

  var pStyle = '';
  for(var i = 0; i < pArray.length; i++){
    var el = pArray[i];
    var text = el.getText();

    // Blank / whitespace-only line: remove it (gaps are margins now). Page breaks kept.
    if(!text || text.trim() === ''){ removeIfBlank(el); continue; }

    // Leave manually-centered paragraphs alone (e.g. a title the user centered).
    if(el.getAlignment() === CENTER){ continue; }

    var upper = text.toUpperCase();

    // SCENE — INT./EXT. (and combined forms), optionally preceded by a scene number.
    if(/^(INT|EXT|EST|I\/E|E\/I)[\.\/]/.test(upper) ||
       (!isNaN(parseInt(text.charAt(0))) && (upper.indexOf('INT.') > -1 || upper.indexOf('EXT.') > -1))){
      var header = upper.replace(/^\s*\d+[\.\):\t ]+/, '');                 // strip old number, keep full prefix
      var headerText = sceneNumbers ? (sceneNum++ + '\t' + header) : header;
      if(headerText !== text){ el.setText(headerText); }
      apply(el, sceneNumbers ? sceneWithNumbers : scene);
      pStyle = 'scene';
      continue;
    }

    // PARENTHETICAL — a line that is WHOLLY "(...)". Inline "(beat) line" is left for DIALOGUE.
    var trimmed = text.trim();
    if(trimmed.charAt(0) === '(' && trimmed.charAt(trimmed.length - 1) === ')'){
      apply(el, parenthetical);
      pStyle = 'parenthetical';
      continue;
    }

    // CENTERED — ">text<"
    if(text.charAt(0) === '>' && text.charAt(text.length - 1) === '<'){
      el.setText(text.substring(1, text.length - 1));
      apply(el, centered);
      pStyle = 'action';
      continue;
    }

    // TRANSITION — ends in " to:", " in:", "out:" (before CHARACTER so "CUT TO:" isn't a name).
    var tail = text.substring(text.length - 4).toLowerCase();
    if(tail === ' to:' || tail === ' in:' || tail === 'out:'){
      apply(el, transition);
      pStyle = 'transition';
      continue;
    }

    // CHARACTER (before DIALOGUE so a real cue is never swallowed). All-caps name, not
    // ending in punctuation, with a "speakable" next line. Also expands shorthands.
    var bracket = upper.indexOf('(');                                       // strip (V.O.)/(CONT'D) for the name
    var nameOnly = bracket > -1 ? upper.substring(0, bracket - 1) : upper;
    var brackets = bracket > -1 ? upper.substring(bracket - 1) : '';
    var expanded = expandShorthand(el, nameOnly, brackets);
    if(expanded !== null){ text = expanded; upper = text.toUpperCase(); }
    if(upper === text){
      var last = upper.slice(-1);
      var endsPunct = endPunctuationMeansNotChar && (last === '.' || last === '!' || last === '?' || last === '-');
      if(!endsPunct && speakableFollows(i)){
        apply(el, character);
        pStyle = 'character';
        if(charList.length < charLimit){ addToCharList(nameOnly); }
        continue;
      }
    }

    // DIALOGUE — directly follows a cue or a parenthetical (keyed only off the prior
    // element, so re-formatting is self-correcting and idempotent).
    if(pStyle === 'character' || pStyle === 'parenthetical'){
      apply(el, dialogue);
      pStyle = 'dialogue';
      continue;
    }

    // ACTION — everything else.
    apply(el, action);
    pStyle = 'action';
  }

  setCharsToStorage(charList);
  return charList;
}

// Format just the selection / the paragraph under the cursor — fast, incremental.
function formatSelection(sceneNumbers, autoFontsMargins, endPunctuationMeansNotChar){
  return convert('selection', sceneNumbers, autoFontsMargins, endPunctuationMeansNotChar);
}

// Returns the paragraphs to format. For 'selection', seeds context with the line above.
function collectParagraphs(type){
  if(type === 'whole'){ return body.getParagraphs(); }

  var els, selection = doc.getSelection();
  if(selection){
    els = selection.getSelectedElements().map(function(r){
      var el = r.getElement();
      // promote a fully-selected text run to its paragraph
      if(typeof el.getText === 'function' && el.getParent() && el.getText() === el.getParent().getText()){
        return el.getParent();
      }
      return el;
    });
  } else {
    var cursor = doc.getCursor();
    if(!cursor){
      DocumentApp.getUi().alert("No Selection",
        'Put your cursor in a line (or highlight a section) to format, or switch to "Whole" under Options.',
        DocumentApp.getUi().ButtonSet.OK);
      return null;
    }
    var el = cursor.getElement();
    while(el && el.getType && el.getType() !== DocumentApp.ElementType.PARAGRAPH){ el = el.getParent(); }
    els = [el];
  }
  els.unshift(elAbove(els[0]));   // the line above seeds pStyle + the speakable look-ahead
  return els;
}

// Remove a blank/whitespace paragraph (but keep page-break paras, e.g. a title page).
function removeIfBlank(el){
  var hasPageBreak = false;
  try { hasPageBreak = el.findElement(DocumentApp.ElementType.PAGE_BREAK) !== null; } catch(e){}
  if(!hasPageBreak){
    try { if(body.getNumChildren() > 1){ body.removeChild(el); } } catch(e){}
  }
}

// If an all-caps name matches a stored shorthand, expand it in place; return the new text (else null).
function expandShorthand(el, nameOnly, brackets){
  for(var j = 0; j < charList.length; j++){
    if(nameOnly === charList[j].sh){
      el.setText(charList[j].name + brackets);
      return el.getText();
    }
  }
  return null;
}

// Text of the next non-blank paragraph after index i (or '' if none).
function nextSpeakableText(i){
  for(var k = i + 1; k < pArray.length; k++){
    var t = '';
    try { t = pArray[k].getText(); } catch(e){ t = ''; }
    if(t && t.trim() !== ''){ return t; }
  }
  return '';
}

// True if the line after i looks like spoken/acted prose (has a lowercase letter, is a
// parenthetical, or ends in !/?) — i.e. the all-caps line at i is a NAME, not a heading.
function speakableFollows(i){
  var nxt = nextSpeakableText(i);
  if(!nxt){ return false; }
  return /[a-z]/.test(nxt) || nxt.charAt(0) === '(' || /[!?]$/.test(nxt.trim());
}

// Document-wide page margins + base font (Courier Prime). Per-element font/spacing
// is set in apply(); this covers the page setup and any unformatted/blank paragraphs.
function docSetUp(){
  doc = getActiveDoc();
  var body = doc.getBody();
  var footer = doc.getFooter() || doc.addFooter();

  body.setMarginTop(72);     // 1.0"
  body.setMarginRight(72);   // 1.0"
  body.setMarginBottom(72);  // 1.0"
  body.setMarginLeft(108);   // 1.5"

  var style = {};
  style[DocumentApp.Attribute.FONT_FAMILY] = screenplayFont;
  style[DocumentApp.Attribute.FONT_SIZE] = 12;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  body.setAttributes(style);

  style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = RIGHT;
  footer.setAttributes(style);
  try { footer.editAsText().setAttributes(style); } catch(e){}
}

// The element 'num' positions above el (defaults to 1, clamped at the top).
function elAbove(el, num){
  var n = num || 1, idx = 0;
  try { idx = doc.getBody().getChildIndex(el); } catch(e){}
  var above = idx - n;
  if(above < 0){ above = 0; }
  try { return doc.getBody().getChild(above); }
  catch(e){ return doc.getBody().getChild(0); }
}

// ---- Per-document storage (characters + sidebar settings) ------------------
function setCharsToStorage(chars){
  try {
    var props = PropertiesService.getDocumentProperties();
    if(!props){ return; }   // no bound document (e.g. running tests)
    props.setProperty('chars', JSON.stringify(chars));
  } catch(e){}
}

function getCharsFromStorage(){
  try {
    var props = PropertiesService.getDocumentProperties();
    if(!props){ return []; }
    return JSON.parse(props.getProperty('chars')) || [];
  } catch(e){ return []; }
}

function getCharsFromDOM(domCharList){
  setCharsToStorage(domCharList);
}

function storeSettings(name, value){
  try {
    var props = PropertiesService.getDocumentProperties();
    if(!props){ return; }
    var settings = JSON.parse(props.getProperty('doc-settings')) || {};
    settings[name] = value;
    props.setProperty('doc-settings', JSON.stringify(settings));
  } catch(e){}
}

function getSettings(){
  try {
    var props = PropertiesService.getDocumentProperties();
    if(!props){ return {}; }
    return JSON.parse(props.getProperty('doc-settings')) || {};
  } catch(e){ return {}; }
}

// Add a character + a minimal unique shorthand to charList (if not already present).
function addToCharList(name){
  for(var i = 0; i < charList.length; i++){
    if(charList[i].name === name){ return; }
  }
  for(var j = 1; j <= name.length; j++){
    var sh = name.substr(0, j), taken = false;
    for(var k = 0; k < charList.length; k++){
      if(charList[k].sh === sh){ taken = true; break; }
    }
    if(!taken){ charList.push({ name: name, sh: sh }); return; }
  }
}
