// Title page: a centered title/subtitle/author block and right-aligned contact
// lines on their own page, inserted before the script. Pro feature.

function addTitlePage(tpInfo){
  var body = DocumentApp.getActiveDocument().getBody();
  var topSpaces = 15;

  body.insertPageBreak(0);
  insertLine(body, tpInfo.contact4, RIGHT);
  insertLine(body, tpInfo.contact3, RIGHT);
  insertLine(body, tpInfo.contact2, RIGHT);
  insertLine(body, tpInfo.contact1, RIGHT);
  for(var i = 0; i < 17; i++){ insertLine(body, '', CENTER); }
  insertLine(body, tpInfo.author, CENTER);
  insertLine(body, '', CENTER);
  insertLine(body, 'by', CENTER);
  insertLine(body, '', CENTER);
  if(tpInfo.subtitle){
    insertLine(body, tpInfo.subtitle, CENTER);
    topSpaces -= 1;
  }
  insertLine(body, tpInfo.title.toUpperCase(), CENTER);
  for(var j = 0; j < topSpaces; j++){ insertLine(body, '', CENTER); }

  storeSettings('tpInfo', tpInfo);
}

function insertLine(body, text, align){
  body.insertParagraph(0, text || '').setAlignment(align);
}

// Remove everything up to (and including) the first page break — or, if the user
// deleted the page break themselves, up to the first scene header.
function removeTitlePage(){
  var body = DocumentApp.getActiveDocument().getBody();
  var cutTo;
  for(var i = 0; i < body.getNumChildren(); i++){
    var p = body.getChild(i);
    var text = p.asParagraph().getText().toUpperCase();
    if(matches([ text.substr(0,4) === 'INT.', text.substr(0,4) === 'EXT.',
                 !isNaN(parseInt(text.substr(0,1))) && (text.indexOf('INT.') > -1 || text.indexOf('EXT.') > -1) ])){
      cutTo = i - 1; // page break gone — stop just before the scene header
      break;
    }
    if(p.findElement(DocumentApp.ElementType.PAGE_BREAK)){ cutTo = i; break; }
  }
  for(var k = 0; k <= cutTo; k++){ body.removeChild(body.getChild(0)); }
}
