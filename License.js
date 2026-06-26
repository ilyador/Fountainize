// Gumroad license check. Free tier gates title page, char shortcuts 6-15, and the
// 800-element limit. License JSON is stored in the user's properties.

var PRODUCT_SUB = 'fountainize-sub';
var PRODUCT_LIFETIME = 'fountainize-lifetime';
var GUMROAD_VERIFY = 'https://api.gumroad.com/v2/licenses/verify';

function isLicenseValid(){
  var stored = PropertiesService.getUserProperties().getProperty('license');
  return stored ? isLicenseCurrent(stored) : false;
}

function isLicenseCurrent(license){
  if(typeof license === 'string'){
    try { license = JSON.parse(license); } catch(e){ return false; }
  }
  if(!license){ return false; }
  // Invalid only if refunded, disputed, or the subscription was cancelled/failed.
  // (Lifetime licenses have no subscription fields — `!= null` leaves them valid.)
  var notValid = license.refunded
    || license.dispute_won
    || license.subscription_cancelled_at != null
    || license.subscription_failed_at != null;
  return !notValid;
}

function getLicense(){
  return PropertiesService.getUserProperties().getProperty('license') || false;
}

function removeLicense(){
  PropertiesService.getUserProperties().deleteProperty('license');
}

// Verify a license key against Gumroad (subscription first, then lifetime) and store it.
function checkLicenseKey(licenseKey){
  return verifyWithProduct(PRODUCT_SUB, licenseKey) || verifyWithProduct(PRODUCT_LIFETIME, licenseKey) || { success: 'false' };
}

function verifyWithProduct(product, licenseKey){
  try {
    var response = UrlFetchApp.fetch(GUMROAD_VERIFY, {
      method: 'post',
      muteHttpExceptions: true,
      payload: { product_permalink: product, license_key: licenseKey }
    });
    var data = JSON.parse(response.getContentText());
    if(data && data.success && data.purchase){
      PropertiesService.getUserProperties().setProperty('license', JSON.stringify(data.purchase));
      return data;
    }
  } catch(e){}
  return null;
}
