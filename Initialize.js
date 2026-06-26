// Add-on lifecycle: menu, install, sidebar, and the "what's new" notice.

function onOpen(e){
  try {
    DocumentApp.getUi().createAddonMenu()
      .addItem('Show sidebar', 'showSidebar')
      .addItem('Format Script', 'convert')
      .addToUi();
  } catch(err){
    DocumentApp.getUi().alert("Issue creating menu",
      "Google prevented us from starting because we don't have the right permissions. Please re-install and allow all permissions. Thanks!",
      DocumentApp.getUi().ButtonSet.OK);
  }

  // Show the changelog once per version bump.
  try {
    var props = PropertiesService.getDocumentProperties();
    if(props && props.getProperty('version') != version){
      DocumentApp.getUi().alert("Fountainize update!",
        "Here's what's new in version " + version + "\n" + changelog[version],
        DocumentApp.getUi().ButtonSet.OK);
      props.setProperty('version', version);
    }
  } catch(err){}
}

function onInstall(e){
  onOpen(e);
}

// Show the 300px sidebar (title reflects license tier).
function showSidebar(){
  var title = isLicenseValid() !== false ? "Fountainize Pro" : "Fountainize (Free)";
  var html = HtmlService.createTemplateFromFile("FountainizeSidebar").evaluate().setTitle(title);
  try { DocumentApp.getUi().showSidebar(html); } catch(e){}
}

var changelog = {
  10: "Added\n- in:/out: now format as transitions\nFixed\n- Issues with elAbove\n- Dialogue in all caps works just fine now\n"
};
