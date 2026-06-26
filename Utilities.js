// Include a JS/CSS file into an HTML template (used by FountainizeSidebar.html).
function include(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// True if any condition in the array is truthy (used by the title-page removal logic).
function matches(conditions){
  for(var i = 0; i < conditions.length; i++){
    if(conditions[i]){ return true; }
  }
  return false;
}
