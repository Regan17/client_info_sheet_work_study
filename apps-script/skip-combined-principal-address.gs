/**
 * Google Apps Script patch — remove duplicate "Current Residential Address" row
 *
 * The HTML form now sends split fields (addrPoBox, addrUnit, addrStreetNo, …) and sets
 * principalAddressUsesSplitFields: true. The generated doc should not also print the
 * legacy combined address row when those fields are present.
 *
 * In your existing Apps Script project (the one deployed at SCRIPT_URL in index.html):
 * 1. Paste the helpers below (or merge into your doc-builder file).
 * 2. Where you append personal-info rows to the Google Doc, skip the combined row when
 *    shouldSkipCombinedPrincipalAddressRow(data, label) is true.
 * 3. Deploy a new version of the web app (same URL if you use "Manage deployments").
 */

function shouldSkipCombinedPrincipalAddressRow(data, label) {
  if (!data) return false;
  if (data.principalAddressUsesSplitFields === true) return isCombinedPrincipalAddressLabel(label);
  return isCombinedPrincipalAddressLabel(label) && hasSplitPrincipalAddressFields(data);
}

function isCombinedPrincipalAddressLabel(label) {
  if (!label) return false;
  var s = String(label).toLowerCase().replace(/\s+/g, ' ');
  return s.indexOf('current residential address') !== -1 && s.indexOf('postal') !== -1;
}

function hasSplitPrincipalAddressFields(data) {
  var keys = [
    'addrPoBox', 'addrUnit', 'addrStreetNo', 'addrStreetName',
    'addrCity', 'addrProvince', 'addrPostal', 'addrCountry'
  ];
  for (var i = 0; i < keys.length; i++) {
    var v = data[keys[i]];
    if (v && String(v).trim()) return true;
  }
  return false;
}

// Example integration inside your doc row loop:
//
// for (var i = 0; i < personalRows.length; i++) {
//   var label = personalRows[i][0];
//   var key = personalRows[i][1];
//   if (shouldSkipCombinedPrincipalAddressRow(data, label)) continue;
//   ...
// }
