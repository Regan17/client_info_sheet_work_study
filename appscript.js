var RESULTS_FOLDER_ID = "1dZr72wia0cgLhMwg-Uk1noR1pMT4QIJr";
var DRAFTS_FOLDER_NAME = "_drafts_autosave";
var STORE = PropertiesService.getScriptProperties();

// ── DRIVE-BACKED DRAFT STORAGE ──
// PropertiesService has a 9KB-per-value / 500KB-total quota that the form blows past
// once a few clients fill out the longer sections. Drafts are stored as JSON files in
// a hidden subfolder of RESULTS_FOLDER_ID instead.
function getDraftsFolder_() {
  var parent = DriveApp.getFolderById(RESULTS_FOLDER_ID);
  var folders = parent.getFoldersByName(DRAFTS_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : parent.createFolder(DRAFTS_FOLDER_NAME);
}

function getDraftFile_(key) {
  var files = getDraftsFolder_().getFilesByName(key + '.json');
  return files.hasNext() ? files.next() : null;
}

function readDraft_(key) {
  var file = getDraftFile_(key);
  if (!file) return null;
  try { return JSON.parse(file.getBlob().getDataAsString()); }
  catch (e) { return null; }
}

function writeDraft_(key, obj) {
  var json = JSON.stringify(obj);
  var file = getDraftFile_(key);
  if (file) file.setContent(json);
  else getDraftsFolder_().createFile(key + '.json', json, 'text/plain');
}

// Read docId from the new Drive draft, falling back to the legacy PropertiesService
// entry so submits made against pre-migration data still update the same doc.
function getStoredDocId_(key) {
  var draft = readDraft_(key);
  if (draft && draft.docId) return draft.docId;
  var legacy = STORE.getProperty(key);
  if (legacy) {
    try { return JSON.parse(legacy).docId || ''; } catch (e) {}
  }
  return '';
}

// Application types that use the same client information sheet (for doc title)
var APPLICATION_TYPE_LABELS = {
  'study_permit_extension': 'Study Permit Extension',
  'pgwp': 'PGWP',
  'dummy_work_permit': 'Dummy Work Permit',
  'visitor_record': 'Visitor Record',
  'visitor_visa': 'Visitor Visa',
  'pnp': 'Manitoba PNP (MPNP)',
  'pnp_based_pr': 'PNP based PR',
  'pr': 'PR (Permanent Residence)',
  'eoi': 'EOI (Expression of Interest)',
  'express_entry': 'Express Entry'
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'save';

    if (action === 'load') {
      return loadClient(data.email, data.applicationType);
    } else {
      return saveClient(data);
    }
  } catch(err) {
    return res({ result:'error', error: err.toString() });
  }
}

// ── LOAD: return saved data for this email + application type ──
function loadClient(email, appType) {
  var key = emailToKey(email, appType);

  // Primary: Drive-backed draft file
  var draft = readDraft_(key);
  if (draft) {
    return res({ result:'success', data: draft.formData, docId: draft.docId });
  }

  // Fallback 1: legacy PropertiesService entry for this exact key
  var stored = STORE.getProperty(key);
  if (stored) {
    var obj = JSON.parse(stored);
    return res({ result:'success', data: obj.formData, docId: obj.docId });
  }

  // Fallback 2: legacy entries saved before per-type keys existed (email-only key).
  // Only used when an appType is supplied so it doesn't shadow newer typed records.
  if (appType) {
    var legacy = STORE.getProperty(emailToKey(email));
    if (legacy) {
      var lobj = JSON.parse(legacy);
      var legacyType = lobj.formData && lobj.formData.applicationType;
      if (!legacyType || legacyType === appType) {
        return res({ result:'success', data: lobj.formData, docId: lobj.docId });
      }
    }
  }
  return res({ result:'not_found' });
}

// ── SAVE / SUBMIT: save data + update Google Doc ──
function saveClient(data) {
  var email = data.email;
  var key = emailToKey(email, data.applicationType);
  var action = data.action || 'save';

  // ── AUTO-SAVE: write the draft JSON file in Drive, no doc work ──
  if (action === 'save') {
    var docId = getStoredDocId_(key);
    writeDraft_(key, { docId: docId, formData: data });
    return res({ result:'success', docId: docId });
  }

  // ── SUBMIT: find folder by passport, create/update doc ──
  var folder = DriveApp.getFolderById(RESULTS_FOLDER_ID);
  var clientName = ((data.givenName||'') + ' ' + (data.lastName||'')).trim() || 'Unknown';
  var passportNo = (data.passport||'').trim();

  // Search for existing folder containing passport number
  var clientFolder = null;
  if (passportNo) {
    var allFolders = folder.getFolders();
    while (allFolders.hasNext()) {
      var f = allFolders.next();
      if (f.getName().indexOf(passportNo) !== -1) {
        clientFolder = f;
        break;
      }
    }
  }

  // No folder found → create one
  if (!clientFolder) {
    var folderName = clientName + (passportNo ? '(' + passportNo + ')' : '');
    clientFolder = folder.createFolder(folderName);
  }

  // Get existing docId from the Drive draft (with PropertiesService fallback)
  var docId = getStoredDocId_(key);

  // Update existing doc or create new one
  var doc;
  var isPnpBasedPr = (data.applicationType === 'pnp_based_pr');
  var isPR = (data.applicationType === 'pr');
  var isPnp = (data.applicationType === 'pnp');
  var isVisitorVisa = (data.applicationType === 'visitor_visa');
  var isEOI = (data.applicationType === 'eoi' || data.applicationType === 'express_entry');
  var fillDocFn = isPR ? fillDocPR
    : (isPnpBasedPr ? fillDocPnpBasedPr
    : (isPnp ? fillDocPNP
    : (isVisitorVisa ? fillDocVisitorVisa
    : (isEOI ? fillDocEOI : fillDoc))));

  if (docId) {
    try {
      doc = DocumentApp.openById(docId);
      var body = doc.getBody();
      body.clear();
      fillDocFn(body, data);
      doc.saveAndClose();
      // Move doc to client folder if it was created elsewhere (e.g. old folder ID)
      try {
        var file = DriveApp.getFileById(docId);
        file.moveTo(clientFolder);
      } catch (moveErr) { /* ignore if already in place */ }
    } catch(e) {
      docId = ''; // doc inaccessible, create fresh
    }
  }

  if (!docId) {
    var docTitle = 'Client Information Sheet';
    var typeLabel = data.applicationType && APPLICATION_TYPE_LABELS[data.applicationType];
    if (typeLabel) docTitle += ' - ' + typeLabel;
    var newDoc = DocumentApp.create(docTitle);
    docId = newDoc.getId();
    fillDocFn(newDoc.getBody(), data);
    newDoc.saveAndClose();
    // Move doc into client folder (moveTo works for both My Drive and Shared Drive)
    try {
      var file = DriveApp.getFileById(docId);
      file.moveTo(clientFolder);
    } catch (moveErr) {
      // If move fails (e.g. Shared Drive restrictions), try add + remove from root
      try {
        var file = DriveApp.getFileById(docId);
        clientFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      } catch (e) {
        throw new Error('Created doc but could not move to folder: ' + (moveErr.message || moveErr.toString()));
      }
    }
  }

  // Save final state to the Drive draft
  writeDraft_(key, { docId: docId, formData: data });

  return res({ result:'success', docId: docId, docUrl: 'https://docs.google.com/document/d/' + docId });
}

function emailToKey(email, appType) {
  var base = 'client_' + email.toLowerCase().replace(/[^a-z0-9]/g, '_');
  if (!appType) return base;
  return base + '__' + String(appType).toLowerCase().replace(/[^a-z0-9]/g, '_');
}

function res(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── FILL DOC ──
// ── Spouse / Children section gating ──
// The form posts every spouse and child field regardless of what the client
// selected, so the doc builders must decide for themselves whether to print
// those sections. A section is printed when the client's own answer calls for
// it, or when the record already holds data there (covers older drafts saved
// before marital status was captured).
function anyValue_(data, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (String(data[keys[i]] || '').trim() !== '') return true;
  }
  return false;
}

function showSpouseSection_(data) {
  var m = String(data.maritalStatus || '').trim().toLowerCase();
  if (m === 'married' || m.indexOf('common-law') === 0 || m.indexOf('common law') === 0) return true;
  // An explicit Single / Divorced / Widowed answer wins, even if the record
  // still carries spouse values the client typed before changing their answer.
  if (m === 'single' || m === 'divorced' || m === 'widowed') return false;
  return anyValue_(data, ['spouseGiven', 'spouseLast', 'spouseDob', 'spouseEmail',
    'spousePhone', 'spouseAddr', 'spouseCanadaStatus', 'spouseEyeColor',
    'spouseHeightCm', 'spouseHeight', 'prevRel', 'prevGiven', 'prevLast',
    'prevDob', 'prevType', 'prevFrom', 'prevTo']);
}

function showChildrenSection_(data) {
  var c = String(data.hasChildren || '').trim().toLowerCase();
  if (c === 'yes') return true;
  if (c === 'no') return false;
  return anyValue_(data, ['child1', 'child2', 'child3', 'childExtra']);
}

function fillDoc(body, data) {

  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }

  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }

  function addSubHeader(body, title) {
    var p = body.appendParagraph(title);
    p.setBold(true).setItalic(true);
    p.setSpacingBefore(8).setSpacingAfter(4);
  }

  // Title
  var title = body.appendParagraph('INFORMATION SHEET');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);

  // Section 1 — Table
  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var isPgwpDoc = data.applicationType === 'pgwp';
  var tableData = [
    ['Given Name', data.givenName||'—'],
    ['Last Name', data.lastName||'—']
  ];
  tableData.push(
    ['Date of Birth (DD/MM/YYYY)', data.dob||'—'],
    ['Email ID', data.email||'—'],
    ['Phone Number', data.phone||'—'],
    ['P.O. Box', data.addrPoBox||'—'],
    ['Apt / Unit Number', data.addrUnit||'—'],
    ['Street Number', data.addrStreetNo||'—'],
    ['Street Name', data.addrStreetName||'—'],
    ['City / Town', data.addrCity||'—'],
    ['Province', data.addrProvince||'—'],
    ['Postal Code', data.addrPostal||'—'],
    ['Country or Territory', data.addrCountry||'—'],
    ['Marital Status', data.maritalStatus||'—'],
    ['Date of Marriage (DD/MM/YYYY)', data.dom||'—'],
    ['Do you have any children?', data.hasChildren||'—'],
    ['Native Language (Mother Tongue)', data.nativeLang||'—']
  );
  tableData.push(['Status in Canada (current)', data.canadaStatus||'—']);
  tableData.push(['Passport No.', data.passport||'—']);
  tableData.push(['UCI No.', data.uci||'—']);
  var table = body.appendTable(tableData);
  table.setBorderWidth(1);
  for (var r=0;r<tableData.length;r++) {
    table.getCell(r,0).setWidth(220).getChild(0).asParagraph().editAsText().setBold(true);
    table.getCell(r,0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    table.getCell(r,1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
  }

  // Section 2 — Spouse
  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    addField(body,'Given Name',data.spouseGiven);
    addField(body,'Last Name',data.spouseLast);
    addField(body,'Date of Birth (DD/MM/YYYY)',data.spouseDob);
    addField(body,'Email ID',data.spouseEmail);
    addField(body,'Phone Number',data.spousePhone);
    addField(body,'Current Address',data.spouseAddr);
    addField(body,'Have you been married or in a common-law relationship before your current marriage?',data.prevRel);
    if(data.prevRel==='Yes'){
      addSubHeader(body,'If yes, kindly provide the following details:');
      addField(body,'Given Name',data.prevGiven);
      addField(body,'Last Name',data.prevLast);
      addField(body,'Date of Birth (DD/MM/YYYY)',data.prevDob);
      addField(body,'Type of Relationship',data.prevType);
      addField(body,'From',data.prevFrom);
      addField(body,'To',data.prevTo);
    }
  }

  // Section 3 — Children
  if (showChildrenSection_(data)) {
    addSectionHeader(body,'🔹 CHILDREN INFO (If Any)');
    addField(body,'1',data.child1);
    addField(body,'2',data.child2);
    addField(body,'3',data.child3);
    if(data.childExtra) addField(body,'Additional',data.childExtra);
  }

  // Section 4 — Other Country
  addSectionHeader(body,'🌍 OTHER COUNTRY RESIDENCY (6+ months)');
  addField(body,'Have you lived in any other country for 6+ months (not Canada or home country)?',data.otherCountry);
  if(data.otherCountry==='Yes'){
    addField(body,'From (MM/YYYY)',data.ocFrom);
    addField(body,'To (MM/YYYY)',data.ocTo);
    addField(body,'Country Name',data.ocCountry);
    addField(body,'Status (e.g., Work/Study Visa, PR)',data.ocStatus);
    addField(body,'Purpose of Stay',data.ocPurpose);
  }

  // Section 5 — Travel
  addSectionHeader(body,'✈️ TRAVEL HISTORY');
  addField(body,'First Entry to Canada (Date)',data.firstEntryDate);
  addField(body,'First Entry to Canada (Airport where first landed)',data.firstEntryPort);
  addField(body,'Recent Entry to Canada (Date)',data.recentEntryDate);
  addField(body,'Recent Entry to Canada (Place where first landed)',data.recentEntryPort);

  // Section 6 — Education
  addSectionHeader(body,'🎓 EDUCATION');
  body.appendParagraph(isPgwpDoc
    ? 'All post-secondary programs — completed or not, inside or outside of Canada.'
    : 'Grade 12 and all post-secondary programs — completed or not, inside or outside of Canada.'
  ).setItalic(true).setSpacingAfter(6);

  // Grade 12
  if (!isPgwpDoc) {
    body.appendParagraph('Grade 12 (High School)').setBold(true);
    addField(body,'From (MM/YYYY)',data.g12from);
    addField(body,'To (MM/YYYY)',data.g12to);
    addField(body,'Program / Stream',data.g12prog);
    addField(body,'School Name',data.g12inst);
    addField(body,'City',data.g12city);
    body.appendParagraph('').setSpacingAfter(4);
  }

  // Post-secondary
  var eduEntries = (data.education||'').split('\n').filter(function(e){
    if(!e.trim()) return false;
    var c=e.replace(/^Entry \d+:\s*/,'').trim();
    return c.replace(/to\s*\|\s*\|\s*\|/,'').replace(/\|/g,'').trim()!=='';
  });
  if(eduEntries.length>0){
    addSubHeader(body,'Post-Secondary Education');
    eduEntries.forEach(function(entry,idx){
      var parts=entry.replace(/^Entry \d+:\s*/,'').split(' | ');
      var dates=(parts[0]||'').split(' to ');
      body.appendParagraph('Education '+(idx+1)).setBold(true);
      addField(body,'From (MM/YYYY)',dates[0]?dates[0].trim():'—');
      addField(body,'To (MM/YYYY)',dates[1]?dates[1].trim():'—');
      addField(body,'Program Name',parts[1]||'—');
      addField(body,'Institute Name',parts[2]||'—');
      addField(body,'Campus City',parts[3]||'—');
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 7 — Work
  addSectionHeader(body,'💼 WORK HISTORY (All works done on SIN in Canada and all works done outside of Canada)');
  var workEntries=(data.work||'').split('\n').filter(function(e){
    if(!e.trim()) return false;
    var c=e.replace(/^Job \d+:\s*/,'').trim();
    return c.replace(/\|/g,'').trim()!=='';
  });
  if(workEntries.length===0){body.appendParagraph('—');}
  else {
    workEntries.forEach(function(entry,idx){
      var parts=entry.replace(/^Job \d+:\s*/,'').split(' | ');
      var dates=(parts[0]||'').split(' to ');
      var isNewFormat = parts.length >= 6;
      body.appendParagraph((idx===0?'1️⃣ Current employment details:':(idx+1)+'️⃣')).setBold(true);
      addField(body,'From (DD/MM/YYYY)',dates[0]?dates[0].trim():'—');
      addField(body,'To (DD/MM/YYYY)',dates[1]?dates[1].trim():'—');
      addField(body,'Job Title',parts[1]||'—');
      if(isNewFormat){
        addField(body,'Full-time or Part-time',parts[2]||'—');
        addField(body,'Hours per week',parts[3]||'—');
        addField(body,'Employer / Company Name',parts[4]||'—');
        addField(body,'Full Employer / Work Location Address',parts[5]||'—');
      } else {
        addField(body,'Employer / Company Name',parts[2]||'—');
        addField(body,'Full Employer / Work Location Address',parts[3]||'—');
      }
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 8 — Financial (not collected for PGWP)
  if (!isPgwpDoc) {
    addSectionHeader(body,'💰 Financial Questions');
    addField(body,'How much do you have in savings or checking accounts?',data.savings);
    addField(body,'If you do not have a job, who and how are you paying for the expenses?',data.expenses);
  }

  // Section 9 — Application History
  addSectionHeader(body,'🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  body.appendParagraph('List all the applications ever applied to IRCC: Study permit, work permit, extensions of all kinds, TRV, visitor.').setItalic(true).setSpacingAfter(6);
  var appRows = data.appHistoryRows || [];
  var appHeaderRow = ['Type of application','Result (Approved/denied)','Year of result','Destination in Canada where received'];
  if(appRows.length===0){
    body.appendTable([appHeaderRow]).setBorderWidth(1);
  } else {
    var appTableData = [appHeaderRow].concat(appRows.map(function(r){
      return [r.type||'—',r.result||'—',r.year||'—',r.destination||'—'];
    }));
    body.appendTable(appTableData).setBorderWidth(1);
  }
  body.appendParagraph('').setSpacingAfter(8);
  addField(body,'Have you ever been refused by IRCC or (US, AUS, NZ or any other country)?',data.refused);
  if(data.refused==='Yes'){
    addField(body,'Country',data.refCountry);
    addField(body,'Result Date (MM/YYYY)',data.refDate);
    addField(body,'Type of Application (Visit / Work / PR)',data.refType);
  }
  addField(body,'Have you ever committed, been arrested for or been charged with or convicted of any criminal offense in any country?',data.criminalRecord);
  if(data.criminalRecord==='Yes'){
    addField(body,'Details',data.criminalInfo);
  }

  // Family Information (conditional — not for PGWP)
  if(data.applicationType !== 'pgwp' && data.showFamilyInfo==='Yes' && data.familyMembers && data.familyMembers.length>0){
    addSectionHeader(body,'👪 FAMILY INFORMATION');
    body.appendParagraph('Family members (parents, siblings, spouse, children). If deceased, date of death in address field.').setItalic(true).setSpacingAfter(6);
    var famHeaders = ['Full Name','Date of Birth','Place of Birth','Marital Status','Relationship','Occupation','Current Address'];
    var famTableData = [famHeaders].concat(data.familyMembers.map(function(m){
      return [m.fullName||'—',m.dob||'—',m.placeOfBirth||'—',m.maritalStatus||'—',m.relationship||'—',m.occupation||'—',m.currentAddress||'—'];
    }));
    body.appendTable(famTableData).setBorderWidth(1);
    body.appendParagraph('').setSpacingAfter(8);
  }

  // Staff Notes
  addSectionHeader(body,'🗒️ Staff Notes / Observations');
  body.appendParagraph('To be filled by our staff only — not applicant').setItalic(true).setSpacingAfter(8);
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client\'s UCI number yourself.').setBold(true).setSpacingAfter(8);
  if (isPgwpDoc) {
    body.appendParagraph('📌 PGWP: Date of Birth and Status in Canada are not asked of PGWP clients, so those rows are blank above. Please fill them in yourself from the file on record.').setBold(true).setSpacingAfter(8);
  }
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  [
    'Current Address: Kindly confirm with the client whether they reside in an apartment building or a house.',
    'Email Address: Please confirm with the client that the email address mentioned above is accurate.',
    'IRCC Refusals and Approvals: Kindly confirm with the client the details of previous application approvals and refusals.',
    'Work History: Please confirm with the client whether they have worked in any other occupation, even for one day, apart from the jobs already listed.',
    'Experience Outside Canada: Have you confirmed whether the client has worked outside Canada, or whether this has been declared in any previous applications to IRCC?',
    'Policy and Procedure: Did you read the policy and any update related to the type of application?',
    'Documents Checklist: Did you make sure that all the documents in the checklist are uploaded and marked with ✅? If anything is missing, please explain the reason.',
    'Passport Expiry: Have you checked if the passport expiry date is more than 3 years? If not, kindly highlight in this section.'
  ].forEach(function(n,i){
    body.appendParagraph((i+1)+'.  '+n).setSpacingAfter(6);
  });
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}

// ── EOI / Express Entry: same as fillDoc but with Relatives in Canada + Miscellaneous Questions ──
function fillDocEOI(body, data) {
  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }
  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }
  function addSubHeader(body, title) {
    var p = body.appendParagraph(title);
    p.setBold(true).setItalic(true);
    p.setSpacingBefore(8).setSpacingAfter(4);
  }

  // Title
  var title = body.appendParagraph('INFORMATION SHEET');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);
  body.appendParagraph('[' + (APPLICATION_TYPE_LABELS[data.applicationType]||data.applicationType||'') + ']').setItalic(true).setSpacingAfter(8);

  // Section 1 — Principal Applicant Table
  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var tableData = [
    ['Given Name', data.givenName||'—'],
    ['Last Name', data.lastName||'—'],
    ['Date of Birth (DD/MM/YYYY)', data.dob||'—'],
    ['Email ID', data.email||'—'],
    ['Phone Number', data.phone||'—'],
    ['Current Residential Address (full with Postal Code)', data.address||'—'],
    ['Marital Status', data.maritalStatus||'—'],
    ['Date of Marriage (DD/MM/YYYY)', data.dom||'—'],
    ['Native Language (Mother Tongue)', data.nativeLang||'—'],
    ['Status in Canada (current)', data.canadaStatus||'—'],
    ['Passport No.', data.passport||'—'],
    ['UCI No.', data.uci||'—']
  ];
  var table = body.appendTable(tableData);
  table.setBorderWidth(1);
  for (var r=0;r<tableData.length;r++) {
    table.getCell(r,0).setWidth(220).getChild(0).asParagraph().editAsText().setBold(true);
    table.getCell(r,0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    table.getCell(r,1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
  }

  // Section 2 — Spouse
  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    addField(body,'Given Name',data.spouseGiven);
    addField(body,'Last Name',data.spouseLast);
    addField(body,'Date of Birth (DD/MM/YYYY)',data.spouseDob);
    addField(body,'Email ID',data.spouseEmail);
    addField(body,'Phone Number',data.spousePhone);
    addField(body,'Current Address',data.spouseAddr);
    addField(body,'Have you been married or in a common-law relationship before your current marriage?',data.prevRel);
    if(data.prevRel==='Yes'){
      addSubHeader(body,'If yes, kindly provide the following details:');
      addField(body,'Given Name',data.prevGiven);
      addField(body,'Last Name',data.prevLast);
      addField(body,'Date of Birth (DD/MM/YYYY)',data.prevDob);
      addField(body,'Type of Relationship',data.prevType);
      addField(body,'From',data.prevFrom);
      addField(body,'To',data.prevTo);
    }
  }

  // Section 3 — Children
  if (showChildrenSection_(data)) {
    addSectionHeader(body,'🔹 CHILDREN INFO (If Any)');
    addField(body,'1',data.child1);
    addField(body,'2',data.child2);
    addField(body,'3',data.child3);
    if(data.childExtra) addField(body,'Additional',data.childExtra);
  }

  // Section 4 — Other Country
  addSectionHeader(body,'🌍 OTHER COUNTRY RESIDENCY (6+ months)');
  addField(body,'Have you lived in any other country for 6+ months (not Canada or home country)?',data.otherCountry);
  if(data.otherCountry==='Yes'){
    addField(body,'From (MM/YYYY)',data.ocFrom);
    addField(body,'To (MM/YYYY)',data.ocTo);
    addField(body,'Country Name',data.ocCountry);
    addField(body,'Status (e.g., Work/Study Visa, PR)',data.ocStatus);
    addField(body,'Purpose of Stay',data.ocPurpose);
  }

  // Section 5 — Travel
  addSectionHeader(body,'✈️ TRAVEL HISTORY');
  addField(body,'First Entry to Canada (Date)',data.firstEntryDate);
  addField(body,'First Entry to Canada (Airport where first landed)',data.firstEntryPort);
  addField(body,'Recent Entry to Canada (Date)',data.recentEntryDate);
  addField(body,'Recent Entry to Canada (Place where first landed)',data.recentEntryPort);

  // Section 6 — Education
  addSectionHeader(body,'🎓 EDUCATION');
  body.appendParagraph('Grade 12 and all post-secondary programs — completed or not, inside or outside of Canada.').setItalic(true).setSpacingAfter(6);
  body.appendParagraph('Grade 12 (High School)').setBold(true);
  addField(body,'From (MM/YYYY)',data.g12from);
  addField(body,'To (MM/YYYY)',data.g12to);
  addField(body,'Program / Stream',data.g12prog);
  addField(body,'School Name',data.g12inst);
  addField(body,'City',data.g12city);
  body.appendParagraph('').setSpacingAfter(4);
  var eduEntries = (data.education||'').split('\n').filter(function(e){
    if(!e.trim()) return false;
    var c=e.replace(/^Entry \d+:\s*/,'').trim();
    return c.replace(/to\s*\|\s*\|\s*\|/,'').replace(/\|/g,'').trim()!=='';
  });
  if(eduEntries.length>0){
    addSubHeader(body,'Post-Secondary Education');
    eduEntries.forEach(function(entry,idx){
      var parts=entry.replace(/^Entry \d+:\s*/,'').split(' | ');
      var dates=(parts[0]||'').split(' to ');
      body.appendParagraph('Education '+(idx+1)).setBold(true);
      addField(body,'From (MM/YYYY)',dates[0]?dates[0].trim():'—');
      addField(body,'To (MM/YYYY)',dates[1]?dates[1].trim():'—');
      addField(body,'Program Name',parts[1]||'—');
      addField(body,'Institute Name',parts[2]||'—');
      addField(body,'Campus City',parts[3]||'—');
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 7 — Work
  addSectionHeader(body,'💼 WORK HISTORY (All works done on SIN in Canada and all works done outside of Canada)');
  var workEntries=(data.work||'').split('\n').filter(function(e){
    if(!e.trim()) return false;
    var c=e.replace(/^Job \d+:\s*/,'').trim();
    return c.replace(/\|/g,'').trim()!=='';
  });
  if(workEntries.length===0){body.appendParagraph('—');}
  else {
    workEntries.forEach(function(entry,idx){
      var parts=entry.replace(/^Job \d+:\s*/,'').split(' | ');
      var dates=(parts[0]||'').split(' to ');
      var isNewFormat = parts.length >= 6;
      body.appendParagraph((idx===0?'1️⃣ Current employment details:':(idx+1)+'️⃣')).setBold(true);
      addField(body,'From (DD/MM/YYYY)',dates[0]?dates[0].trim():'—');
      addField(body,'To (DD/MM/YYYY)',dates[1]?dates[1].trim():'—');
      addField(body,'Job Title',parts[1]||'—');
      if(isNewFormat){
        addField(body,'Full-time or Part-time',parts[2]||'—');
        addField(body,'Hours per week',parts[3]||'—');
        addField(body,'Employer / Company Name',parts[4]||'—');
        addField(body,'Full Employer / Work Location Address',parts[5]||'—');
      } else {
        addField(body,'Employer / Company Name',parts[2]||'—');
        addField(body,'Full Employer / Work Location Address',parts[3]||'—');
      }
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 8 — Application History
  addSectionHeader(body,'🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  body.appendParagraph('List all the applications ever applied to IRCC: Study permit, work permit, extensions of all kinds, TRV, visitor.').setItalic(true).setSpacingAfter(6);
  var appRows = data.appHistoryRows || [];
  var appHeaderRow = ['Type of application','Result (Approved/denied)','Year of result','Destination in Canada where received'];
  if(appRows.length===0){
    body.appendTable([appHeaderRow]).setBorderWidth(1);
  } else {
    body.appendTable([appHeaderRow].concat(appRows.map(function(r){
      return [r.type||'—',r.result||'—',r.year||'—',r.destination||'—'];
    }))).setBorderWidth(1);
  }
  body.appendParagraph('').setSpacingAfter(8);
  addField(body,'Have you ever been refused by IRCC or (US, AUS, NZ or any other country)?',data.refused);
  if(data.refused==='Yes'){
    addField(body,'Country',data.refCountry);
    addField(body,'Result Date (MM/YYYY)',data.refDate);
    addField(body,'Type of Application (Visit / Work / PR)',data.refType);
  }

  // Section 9 — Financial
  addSectionHeader(body,'💰 Financial Questions');
  addField(body,'How much do you have in savings or checking accounts?',data.savings);
  addField(body,'If you do not have a job, who and how are you paying for the expenses?',data.expenses);

  // Family Information (conditional)
  if(data.showFamilyInfo==='Yes' && data.familyMembers && data.familyMembers.length>0){
    addSectionHeader(body,'👪 FAMILY INFORMATION');
    body.appendParagraph('Family members (parents, siblings, spouse, children). If deceased, date of death in address field.').setItalic(true).setSpacingAfter(6);
    var famHeaders = ['Full Name','Date of Birth','Place of Birth','Marital Status','Relationship','Occupation','Current Address'];
    body.appendTable([famHeaders].concat(data.familyMembers.map(function(m){
      return [m.fullName||'—',m.dob||'—',m.placeOfBirth||'—',m.maritalStatus||'—',m.relationship||'—',m.occupation||'—',m.currentAddress||'—'];
    }))).setBorderWidth(1);
    body.appendParagraph('').setSpacingAfter(8);
  }

  // Section 10 — Relatives in Canada (EOI / Express Entry specific)
  addSectionHeader(body,'👪 Relatives in Canada');
  addField(body,'Do you or your spouse have a close relative (PR or citizen) in Manitoba for the past 1 year?',data.relMB);
  if(data.relMB==='Yes' && data.relativesMB && data.relativesMB.length>0){
    data.relativesMB.forEach(function(rel,idx){
      body.appendParagraph('Manitoba Relative '+(idx+1)).setBold(true);
      addField(body,'Name',rel.name);
      addField(body,'Current Address',rel.address);
      addField(body,'Relationship to Applicant',rel.relationship);
      addField(body,'Status in Canada',rel.status);
      addField(body,'Date started living in Manitoba',rel.date);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }
  body.appendParagraph('').setSpacingAfter(4);
  addField(body,'Do you or your spouse have a close relative in any other Canadian province except Manitoba?',data.relOtherProv);
  if(data.relOtherProv==='Yes' && data.relativesOtherProv && data.relativesOtherProv.length>0){
    data.relativesOtherProv.forEach(function(rel,idx){
      body.appendParagraph('Other Province Relative '+(idx+1)).setBold(true);
      addField(body,'Name',rel.name);
      addField(body,'Current Address',rel.address);
      addField(body,'Relationship to Applicant',rel.relationship);
      addField(body,'Status in Canada',rel.status);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 11 — Miscellaneous Questions (EOI / Express Entry specific)
  addSectionHeader(body,'📋 Miscellaneous Questions');
  var miscQuestions = [
    ['Are you applying for the EOI for the first time?', data.miscQ1],
    ['Have you previously been refused by the MPNP within the last six months?', data.miscQ2],
    ['Have you or your spouse ever visited or lived, studied or worked in another Canadian Province or Territory?', data.miscQ3],
    ['Do you have an active Express Entry profile submitted in the past 12 months?', data.miscQ4],
    ['Have you been working in Manitoba full-time for the same employer for the last six months?', data.miscQ5],
    ['Have you received a long-term, full-time job offer from your employer in Manitoba?', data.miscQ6],
    ['Have you ever applied for Express Entry before?', data.miscQ7],
    ['Did you complete at least 50% of the Canadian study program through in-person learning?', data.miscQ8],
    ['Did you study in Canada for at least 8 months?', data.miscQ9],
    ['Did you study full-time for at least 8 months?', data.miscQ10]
  ];
  miscQuestions.forEach(function(q, i) {
    addField(body, (i+1) + '. ' + q[0], q[1]);
  });
  if(data.miscQ3==='Yes' && data.miscQ3Where){
    addField(body,'If yes, where?',data.miscQ3Where);
  }

  // Staff Notes
  addSectionHeader(body,'🗒️ Staff Notes / Observations');
  body.appendParagraph('To be filled by our staff only — not applicant').setItalic(true).setSpacingAfter(8);
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client\'s UCI number yourself.').setBold(true).setSpacingAfter(8);
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  [
    'Current Address: Kindly confirm with the client whether they reside in an apartment building or a house.',
    'Email Address: Please confirm with the client that the email address mentioned above is accurate.',
    'IRCC Refusals and Approvals: Kindly confirm with the client the details of previous application approvals and refusals.',
    'Work History: Please confirm with the client whether they have worked in any other occupation, even for one day, apart from the jobs already listed.',
    'Experience Outside Canada: Have you confirmed whether the client has worked outside Canada, or whether this has been declared in any previous applications to IRCC?',
    'Policy and Procedure: Did you read the policy and any update related to the type of application?',
    'Documents Checklist: Did you make sure that all the documents in the checklist are uploaded and marked with ✅? If anything is missing, please explain the reason.',
    'Passport Expiry: Have you checked if the passport expiry date is more than 3 years? If not, kindly highlight in this section.'
  ].forEach(function(n,i){
    body.appendParagraph((i+1)+'.  '+n).setSpacingAfter(6);
  });
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}

// ── Manitoba PNP (MPNP): PR-depth sheet without Travel, Personal History,
// Address History or Family Information. Spouse mirrors are limited to
// Application History, Education, Work and Background. ──
function fillDocPNP(body, data) {
  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }
  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }
  function addSubHeader(body, title) {
    var p = body.appendParagraph(title);
    p.setBold(true).setItalic(true);
    p.setSpacingBefore(8).setSpacingAfter(4);
  }
  function addTableFromRows(body, headers, rows, keys) {
    if (!rows || rows.length === 0) {
      body.appendTable([headers]).setBorderWidth(1);
      return;
    }
    var tableData = [headers].concat(rows.map(function(r) {
      return keys.map(function(k) { return r[k] || '—'; });
    }));
    body.appendTable(tableData).setBorderWidth(1);
  }
  // The form always posts one blank row per repeatable block, so drop rows
  // where every field is empty before deciding whether a section has content.
  function nonEmpty(rows) {
    if (!rows || !rows.length) return [];
    return rows.filter(function(r) {
      for (var k in r) {
        if (r.hasOwnProperty(k) && String(r[k] || '').trim() !== '') return true;
      }
      return false;
    });
  }

  var title = body.appendParagraph('INFORMATION SHEET (MANITOBA PNP)');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);

  // Section 1 — Principal Applicant
  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var tableData = [
    ['Given Name', data.givenName||'—'],
    ['Last Name', data.lastName||'—'],
    ['Date of Birth (DD/MM/YYYY)', data.dob||'—'],
    ['Email ID', data.email||'—'],
    ['Phone Number', data.phone||'—'],
    ['Current Residential Address (full with Postal code)', data.address||'—'],
    ['Marital Status', data.maritalStatus||'—'],
    ['Date of Marriage (DD/MM/YYYY)', data.dom||'—'],
    ['Native Language (Mother Tongue)', data.nativeLang||'—'],
    ['Status in Canada (current)', data.canadaStatus||'—'],
    ['Eye Color', data.eyeColor||'—'],
    ['Height (in cm)', data.heightCm||'—'],
    ['Passport No.', data.passport||'—'],
    ['UCI No.', data.uci||'—'],
    ['Have you ever used any other name? (yes/no)', data.otherNameUsed||'—']
  ];
  if (data.otherNameUsed === 'Yes') {
    tableData.push(['Other Name(s) Used', data.otherNames||'—']);
  }
  tableData.push(['Date of last entry to Canada', data.prLastEntryDate||'—']);
  tableData.push(['Place of last entry to Canada', data.prLastEntryPlace||'—']);
  tableData.push(['Have you previously been married or in a common-law relationship? (yes/no)', data.prPrevMarried||'—']);
  if (data.prPrevMarried === 'Yes') {
    tableData.push(['Previous Partner — Given Name', data.prPrevGiven||'—']);
    tableData.push(['Previous Partner — Last Name', data.prPrevLast||'—']);
    tableData.push(['Previous Partner — Date of Birth', data.prPrevDob||'—']);
    tableData.push(['Previous Partner — Type of Relationship', data.prPrevRelType||'—']);
    tableData.push(['Previous Partner — From', data.prPrevFrom||'—']);
    tableData.push(['Previous Partner — To', data.prPrevTo||'—']);
  }
  var table = body.appendTable(tableData);
  table.setBorderWidth(1);
  for (var r = 0; r < tableData.length; r++) {
    table.getCell(r, 0).setWidth(280).getChild(0).asParagraph().editAsText().setBold(true);
    table.getCell(r, 0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    table.getCell(r, 1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
  }
  body.appendParagraph('').setSpacingAfter(8);

  // Section 2 — Spouse
  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    addField(body, 'Given Name', data.spouseGiven);
    addField(body, 'Last Name', data.spouseLast);
    addField(body, 'Date of Birth (DD/MM/YYYY)', data.spouseDob);
    addField(body, 'Email ID', data.spouseEmail);
    addField(body, 'Phone Number', data.spousePhone);
    addField(body, 'Current Address', data.spouseAddr);
    addField(body, 'Status in Canada (if in Canada)', data.spouseCanadaStatus);
    addField(body, 'Eye Color', data.spouseEyeColor);
    addField(body, 'Height (in cm)', data.spouseHeightCm);
    addField(body, 'Have you been married or in a common-law relationship before your current marriage?', data.prevRel);
    if (data.prevRel === 'Yes') {
      addSubHeader(body, 'If yes, kindly provide the following details:');
      addField(body, 'Given Name', data.prevGiven);
      addField(body, 'Last Name', data.prevLast);
      addField(body, 'Date of Birth (DD/MM/YYYY)', data.prevDob);
      addField(body, 'Type of Relationship', data.prevType);
      addField(body, 'From', data.prevFrom);
      addField(body, 'To', data.prevTo);
    }
    body.appendParagraph('').setSpacingAfter(8);
  }

  // Section 3 — Children
  if (showChildrenSection_(data)) {
    addSectionHeader(body, '🔹 CHILDREN INFO (If Any)');
    addField(body, '1', data.child1);
    addField(body, '2', data.child2);
    addField(body, '3', data.child3);
    if (data.childExtra) addField(body, 'Additional', data.childExtra);
  }

  // Section 4 — Other Country Residency
  addSectionHeader(body, '🌍 OTHER COUNTRY RESIDENCY (6+ months)');
  addField(body, 'Have you lived in any other country for 6+ months (not Canada or home country)?', data.otherCountry);
  if (data.otherCountry === 'Yes') {
    addField(body, 'From (MM/YYYY)', data.ocFrom);
    addField(body, 'To (MM/YYYY)', data.ocTo);
    addField(body, 'Country Name', data.ocCountry);
    addField(body, 'Status (e.g., Work/Study Visa, PR)', data.ocStatus);
    addField(body, 'Purpose of Stay', data.ocPurpose);
  }

  // Section 5 — Education
  addSectionHeader(body, '🎓 EDUCATION');
  body.appendParagraph('Grade 12 and all post-secondary programs — completed or not, inside or outside of Canada.').setItalic(true).setSpacingAfter(6);
  body.appendParagraph('Grade 12 (High School)').setBold(true);
  addField(body, 'From (MM/YYYY)', data.g12from);
  addField(body, 'To (MM/YYYY)', data.g12to);
  addField(body, 'Program / Stream', data.g12prog);
  addField(body, 'School Name', data.g12inst);
  addField(body, 'City', data.g12city);
  if (data.g12level) addField(body, 'Level of Education', data.g12level);
  if (data.g12field) addField(body, 'Field of Study', data.g12field);
  body.appendParagraph('').setSpacingAfter(4);

  var eduEntries = (data.education || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Entry \d+:\s*/, '').trim();
    return c.replace(/to\s*\|\s*\|\s*\|/, '').replace(/\|/g, '').trim() !== '';
  });
  var eduExtras = data.educationExtras || [];
  if (eduEntries.length > 0) {
    addSubHeader(body, 'Post-Secondary Education');
    eduEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Entry \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      var extra = eduExtras[idx] || {};
      body.appendParagraph('Education ' + (idx + 1)).setBold(true);
      addField(body, 'From (MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Program Name', parts[1] || '—');
      addField(body, 'Institute Name', parts[2] || '—');
      addField(body, 'Campus City', parts[3] || '—');
      if (extra.level) addField(body, 'Level of Education', extra.level);
      if (extra.field) addField(body, 'Field of Study', extra.field);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 6 — Work History
  addSectionHeader(body, '💼 WORK HISTORY (All works done on SIN in Canada and all works done outside of Canada)');
  var workEntries = (data.work || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Job \d+:\s*/, '').trim();
    return c.replace(/\|/g, '').trim() !== '';
  });
  var workExtras = data.workContactExtras || [];
  if (workEntries.length === 0) { body.appendParagraph('—'); }
  else {
    workEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Job \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      var extra = workExtras[idx] || {};
      body.appendParagraph((idx === 0 ? '1️⃣ Current employment details:' : (idx + 1) + '️⃣')).setBold(true);
      addField(body, 'From (DD/MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (DD/MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Job Title', parts[1] || '—');
      addField(body, 'Full-time or Part-time', parts[2] || '—');
      addField(body, 'Hours per week', parts[3] || '—');
      addField(body, 'Employer / Company Name', parts[4] || '—');
      addField(body, 'Full Work Location Address', parts[5] || '—');
      if (extra.contactName) addField(body, 'Contact Person Name', extra.contactName);
      if (extra.contactPhone) addField(body, 'Contact Person Phone', extra.contactPhone);
      if (extra.contactEmail) addField(body, 'Contact Email', extra.contactEmail);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 7 — Financial
  addSectionHeader(body, '💰 Financial Questions');
  addField(body, 'How much do you have in savings or checking accounts?', data.savings);
  addField(body, 'If you do not have a job, who and how are you paying for the expenses?', data.expenses);

  // Section 8 — Applications History and Background
  addSectionHeader(body, '🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  body.appendParagraph('List all applications submitted to IRCC.').setItalic(true).setSpacingAfter(6);
  var appRows = nonEmpty(data.appHistoryRows);
  var appHeaderRow = ['Type of application', 'Result (Approved/denied)', 'Date of result', 'Destination in Canada', 'Reason for refusal (if known)'];
  if (appRows.length === 0) {
    body.appendTable([appHeaderRow]).setBorderWidth(1);
  } else {
    body.appendTable([appHeaderRow].concat(appRows.map(function(r) {
      return [r.type || '—', r.result || '—', r.date || '—', r.destination || '—', r.reason || '—'];
    }))).setBorderWidth(1);
  }
  body.appendParagraph('').setSpacingAfter(8);
  addField(body, 'Have you ever been refused by IRCC or (US, AUS, NZ or any other country)?', data.refused);
  if (data.refused === 'Yes') {
    addField(body, 'Country', data.refCountry);
    addField(body, 'Result Date (MM/YYYY)', data.refDate);
    addField(body, 'Type of Application (Visit / Work / PR)', data.refType);
  }
  addField(body, 'Have you ever committed, been arrested for or been charged with or convicted of any criminal offense in any country?', data.criminalRecord);
  if (data.criminalRecord === 'Yes') {
    addField(body, 'Details', data.criminalInfo);
  }

  // Section 9 — Background Questions
  addSectionHeader(body, '⚖️ BACKGROUND QUESTIONS');
  addField(body, 'Member of any political/social/youth/student organization, trade unions or professional associations?', data.bgPolitical);
  addField(body, 'Ever held a government position (civil servant, judge, police officer, security organization)?', data.bgGovt);
  addField(body, 'Ever served in military or paramilitary service?', data.bgMilitary);
  body.appendParagraph('').setSpacingAfter(8);

  // Section 10 — Spouse sections (App History, Education, Work, Background)
  var spAppHistory = nonEmpty(data.spouseAppHistory);
  var spEducation = nonEmpty(data.spouseEducation);
  var spWork = nonEmpty(data.spouseWork);
  var hasSpouseBg = !!(data.spBgPolitical || data.spBgGovt || data.spBgMilitary);
  var hasSpouseSections = spAppHistory.length > 0 || spEducation.length > 0 ||
    spWork.length > 0 || hasSpouseBg;

  if (hasSpouseSections) {
    body.appendParagraph('IF APPLYING FOR YOUR SPOUSE ALSO, THEN PROVIDE THE FOLLOWING INFORMATION FOR SPOUSE:').setBold(true).setSpacingBefore(16).setSpacingAfter(8);

    if (spAppHistory.length > 0) {
      addSectionHeader(body, '🗂️ Spouse — Application History');
      addTableFromRows(body,
        ['Type of application', 'Result', 'Date of result', 'Destination in Canada', 'Reason for refusal'],
        spAppHistory,
        ['type', 'result', 'date', 'destination', 'reason']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }

    if (spEducation.length > 0) {
      addSectionHeader(body, '🎓 Spouse — Education History');
      addTableFromRows(body,
        ['From (YYYY-MM)', 'To (YYYY-MM)', 'Institution', 'City and Country', 'Level of Education', 'Field of Study'],
        spEducation,
        ['from', 'to', 'institution', 'city', 'level', 'field']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }

    if (spWork.length > 0) {
      addSectionHeader(body, '💼 Spouse — Work History');
      spWork.forEach(function(w, idx) {
        body.appendParagraph('Spouse Job ' + (idx + 1)).setBold(true);
        addField(body, 'From', w.from);
        addField(body, 'To', w.to);
        addField(body, 'Job Title', w.title);
        addField(body, 'Employer / Company Name', w.employer);
        addField(body, 'Full Work Location Address', w.address);
        if (w.contactName) addField(body, 'Contact Person Name', w.contactName);
        if (w.contactPhone) addField(body, 'Contact Person Phone', w.contactPhone);
        if (w.contactEmail) addField(body, 'Contact Email', w.contactEmail);
        body.appendParagraph('').setSpacingAfter(4);
      });
      body.appendParagraph('').setSpacingAfter(4);
    }

    if (hasSpouseBg) {
      addSectionHeader(body, '⚖️ Spouse — Background Questions');
      addField(body, 'Member of any political/social/youth/student organization?', data.spBgPolitical);
      addField(body, 'Ever held a government position?', data.spBgGovt);
      addField(body, 'Ever served in military or paramilitary service?', data.spBgMilitary);
      body.appendParagraph('').setSpacingAfter(8);
    }
  }

  // Section 11 — Relatives in Canada
  addSectionHeader(body, '👪 Relatives in Canada');
  addField(body, 'Do you have any close relatives in Canada? (yes/no)', data.prRelCan);
  var relatives = nonEmpty(data.prRelativesCanada);
  if (data.prRelCan === 'Yes' && relatives.length > 0) {
    relatives.forEach(function(rel, idx) {
      body.appendParagraph('Relative ' + (idx + 1)).setBold(true);
      addField(body, 'Name', rel.name);
      addField(body, 'Date of Birth', rel.dob);
      addField(body, 'Email ID', rel.email);
      addField(body, 'Phone Number', rel.phone);
      addField(body, 'Relationship to Applicant', rel.relationship);
      addField(body, 'Current Address', rel.address);
      addField(body, 'Status in Canada', rel.status);
      addField(body, 'Date they became PR (if PR or Citizen)', rel.prDate);
      addField(body, 'Date they moved to MB (if PR or Citizen)', rel.mbDate);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Staff Notes
  addSectionHeader(body, '🗒️ Staff Notes / Observations (Staff Use Only)');
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client’s UCI number yourself.').setBold(true).setSpacingAfter(8);
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  var staffChecks = [
    ['Current Address confirmed (apartment or house)?', data.staffAddr],
    ['Email Address confirmed as accurate?', data.staffEmail],
    ['IRCC Refusals and Approvals confirmed?', data.staffIRCC],
    ['Native Language names confirmed (except Punjabi/Hindi)?', data.staffLang],
    ['PCC confirmed from specific departments?', data.staffPCC],
    ['Work History — confirmed no other occupation is missing?', data.staffGaps]
  ];
  staffChecks.forEach(function(item, i) {
    body.appendParagraph((i + 1) + '.  ' + item[0] + '  ' + (item[1] || '—')).setSpacingAfter(6);
  });
  if (data.staffMethod) addField(body, 'Method used to confirm information', data.staffMethod);
  if (data.staffNotes) addField(body, 'Additional Staff Notes', data.staffNotes);
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}

// ── PNP based PR: INFORMATION SHEET (PR) — principal + spouse + family table + app history + education + personal history + address + travel + optional spouse repeat + staff notes ──
function fillDocPnpBasedPr(body, data) {
  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }
  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }
  function addTableFromRows(body, headers, rows, keys) {
    if (!rows || rows.length === 0) {
      var tbl = body.appendTable([headers]);
      tbl.setBorderWidth(1);
      return;
    }
    var tableData = [headers].concat(rows.map(function(r) { return keys.map(function(k) { return r[k] || '—'; }); }));
    var tbl = body.appendTable(tableData);
    tbl.setBorderWidth(1);
  }

  var title = body.appendParagraph('INFORMATION SHEET (PR)');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);

  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var tableData = [
    ['Given Name', data.givenName||'—'],
    ['Last Name', data.lastName||'—'],
    ['Date of Birth (DD/MM/YYYY)', data.dob||'—'],
    ['Email ID', data.email||'—'],
    ['Phone Number', data.phone||'—'],
    ['Current Residential Address (full with Postal code)', data.address||'—'],
    ['Marital Status (Married / Single / Divorced / Widowed)', data.maritalStatus||'—'],
    ['Date of Marriage (if married) (DD/MM/YYYY)', data.dom||'—'],
    ['Native Language (Mother Tongue)', data.nativeLang||'—'],
    ['Status in Canada (current)', data.canadaStatus||'—'],
    ['Eye Color', data.eyeColor||'—'],
    ['Height (in cm)', data.height||'—'],
    ['Passport No.', data.passport||'—'],
    ['UCI No.', data.uci||'—'],
    ['Have you ever used any other name (e.g. nickname, maiden name, alias, etc.)? (yes/no)', data.otherName||'—'],
    ['Date and place of your last entry to Canada (if you never went outside of Canada, then date and place of 1st Entry to Canada)?', ''],
    ['Date (YYYY/MM/DD)', data.lastEntryDate||'—'],
    ['Place', data.lastEntryPlace||'—'],
    ['Have you previously been married or in a common-law relationship? (yes or no)', data.previouslyMarried||'—']
  ];
  var table = body.appendTable(tableData);
  table.setBorderWidth(1);
  for (var r = 0; r < tableData.length; r++) {
    table.getCell(r, 0).setWidth(280).getChild(0).asParagraph().editAsText().setBold(true);
    table.getCell(r, 0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    table.getCell(r, 1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
  }
  body.appendParagraph('').setSpacingAfter(8);

  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    body.appendParagraph('If Married or in a common-law partnership, then provide the following information:').setItalic(true).setSpacingAfter(6);
    addField(body, 'Given Name', data.spouseGiven);
    addField(body, 'Last Name', data.spouseLast);
    addField(body, 'Date of Birth (DD/MM/YYYY)', data.spouseDob);
    addField(body, 'Email ID', data.spouseEmail);
    addField(body, 'Phone Number', data.spousePhone);
    addField(body, 'Current Address', data.spouseAddr);
    addField(body, 'Status in Canada (if in Canada)', data.spouseCanadaStatus);
    addField(body, 'Eye Color', data.spouseEyeColor);
    addField(body, 'Height (in cms)', data.spouseHeight);
    addField(body, 'Have you been married or in a common-law relationship before your current marriage? YES/ NO', data.prevRel);
    if (data.prevRel === 'Yes') {
      addField(body, 'Given Name', data.prevGiven);
      addField(body, 'Last Name', data.prevLast);
      addField(body, 'Date of Birth (DD/MM/YYYY)', data.prevDob);
      addField(body, 'Type of Relationship (married/ common-law)', data.prevType);
      addField(body, 'From', data.prevFrom);
      addField(body, 'To', data.prevTo);
    }
    body.appendParagraph('').setSpacingAfter(8);
  }

  addSectionHeader(body, '👪 Family Information:');
  body.appendParagraph('Provide all the information related to your family members (including spouse, parents, children, siblings): (if deceased, then provide the date of death in the current address section with city and country name)').setItalic(true).setSpacingAfter(6);
  var familyHeaders = ['Full Name', 'Date of Birth (DD/MM/YYY)', 'Place of Birth with country name', 'Marital Status', 'Relationship to the applicant', 'Email address (mandatory)', 'Current address (full with postal code)'];
  var familyKeys = ['fullName', 'dob', 'placeOfBirth', 'maritalStatus', 'relationship', 'email', 'currentAddress'];
  addTableFromRows(body, familyHeaders, data.familyMembers || [], familyKeys);
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  addField(body, '1. Have you previously been refused any type of application or denied entry to Canada or any other country? (Yes/No)', data.refusedApp);
  if (data.refusedApp === 'Yes') {
    body.appendParagraph('If yes, provide the following details:').setSpacingAfter(4);
    addTableFromRows(body, ['Type of application', 'Date of result (DD/MM/YYY)', 'Reason for rejection/refusal (if known)', 'Destination in Canada where received'], data.appHistoryRowsPR || [], ['type', 'dateResult', 'reason', 'destination']);
  }
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '📖 EDUCATION HISTORY (List all the completed and incomplete education programs after Grade 10):');
  body.appendParagraph('Note: If you have applied your PNP through us then, do not rewrite the educational details mentioned previously:').setItalic(true).setSpacingAfter(4);
  addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Name of School or Institution', 'City and Country', 'Level Education', 'Field of Study'], data.educationRowsPR || [], ['from', 'to', 'institution', 'cityCountry', 'level', 'fieldOfStudy']);
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '💼 PERSONAL HISTORY:');
  body.appendParagraph('Provide the details of your personal history since the age of 18 or the past 10 years, whichever is most recent. Under activity, write your occupation or job title if you were working. If you were not working, provide information on what you were doing (For example: unemployed, studying, travelling, retired, in detention etc.)').setItalic(true).setSpacingAfter(2);
  body.appendParagraph('Note: If you applied your PNP through us, then do not rewrite the work details mentioned previously. If you were Outside of Canada or your home country, then indicate your status in that country or territory').setItalic(true).setSpacingAfter(2);
  body.appendParagraph('PLEASE DO NOT LEAVE ANY GAPS IN TIME').setBold(true).setSpacingAfter(6);
  addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Activity (Job title /unemployed /Study / Travel / Retired)', 'City and Country', 'Status in Country', 'Name of company, employer/company, school, facility, as applicable'], data.personalHistoryRows || [], ['from', 'to', 'activity', 'cityCountry', 'statusInCountry', 'nameOfCompany']);
  addField(body, 'Have you ever been a member of any political, social, youth or student organization, trade unions and professional associations? (yes or no)', data.govtOrg);
  addField(body, 'Have you ever held a government position (such as civil servant, judge, police officer, employee in a security organization)? (yes or no)', data.govtPosition);
  addField(body, 'Have you ever served in military or paramilitary service for any country? (yes or no)', data.military);
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '🏠 ADDRESS HISTORY');
  body.appendParagraph('List all the addresses where you have lived since your 18th Birthday or the Past 10 years, whichever is most recent. (Mention all address of inside and outside of Canada)').setItalic(true).setSpacingAfter(2);
  body.appendParagraph('NOTE: DO NOT LEAVE ANY GAPS AND MENTION FULL ADDRESS WITH POSTAL CODES.').setBold(true).setSpacingAfter(6);
  addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Street number and name', 'City or Town', 'Province, State or District', 'Postal code/ Zip code', 'Country or Territory'], data.addressHistoryRows || [], ['from', 'to', 'street', 'city', 'province', 'postalCode', 'country']);
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '✈️ TRAVEL HISTORY');
  body.appendParagraph('List all trips you have taken outside of Canada or your home country in the last 10 years or since the age of 18. Include all trips: tourism, business, training etc.').setItalic(true).setSpacingAfter(2);
  body.appendParagraph('NOTE: LIST ALL THE COUNTRIES VISITED DURING YOUR TRANSIT/LAYOVER WHOSE STAMPS HAVE BEEN MARKED ON YOUR PASSPORT.').setBold(true).setSpacingAfter(6);
  addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Duration', 'Destination (City and Country)', 'Purpose of Visit', 'Provide details (if any)'], data.travelHistoryRowsPR || [], ['from', 'to', 'duration', 'destination', 'purpose', 'details']);
  body.appendParagraph('').setSpacingAfter(8);

  addSectionHeader(body, '💰 Financial Questions :');
  addField(body, 'How much do you have in savings or checking accounts?', data.savings);
  body.appendParagraph('').setSpacingAfter(8);

  if (data.spouseAlso === 'Yes') {
    body.appendParagraph('IF APPLYING FOR YOUR SPOUSE ALSO, THEN PROVIDE THE FOLLOWING INFORMATION FOR SPOUSE:').setBold(true).setSpacingAfter(8);
    addSectionHeader(body, '👪 Family Information:');
    body.appendParagraph('Provide all the information related to your family members (including spouse, parents, children, siblings): (if deceased, then provide the date of death in the current address section with city and country name)').setItalic(true).setSpacingAfter(6);
    addTableFromRows(body, familyHeaders, data.spouseFamilyMembers || [], familyKeys);
    body.appendParagraph('').setSpacingAfter(8);
    addSectionHeader(body, '🗂️ APPLICATIONS HISTORY AND BACKGROUND');
    addField(body, '1. Have you previously been refused or denied entry to any country or have been refused for any type of application within Canada? (yes or no)', data.spouseRefusedApp);
    if (data.spouseRefusedApp === 'Yes') {
      body.appendParagraph('If yes, provide the following details:').setSpacingAfter(4);
      addTableFromRows(body, ['Type of application', 'Date of result', 'Reason for rejection/refusal (if known)', 'Destination in Canada where received'], data.spouseAppHistoryRowsPR || [], ['type', 'dateResult', 'reason', 'destination']);
    }
    body.appendParagraph('').setSpacingAfter(8);
    addSectionHeader(body, '📖 EDUCATION HISTORY (List all the completed and incomplete education programs after Grade 10):');
    body.appendParagraph('Note: If you have applied your PNP through us then, do not rewrite the educational details mentioned previously:').setItalic(true).setSpacingAfter(4);
    addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Name of School or Institution', 'City and Country', 'Level Education', 'Field of Study'], data.spouseEducationRowsPR || [], ['from', 'to', 'institution', 'cityCountry', 'level', 'fieldOfStudy']);
    body.appendParagraph('').setSpacingAfter(8);
    addSectionHeader(body, '💼 PERSONAL HISTORY:');
    body.appendParagraph('Provide the details of your personal history since the age of 18 or the past 10 years, whichever is most recent. Under activity, write your occupation or job title if you were working. If you were not working, provide information on what you were doing (For example: unemployed, studying, travelling, retired, in detention etc.)').setItalic(true).setSpacingAfter(2);
    body.appendParagraph('Note: If you applied your PNP through us, then do not rewrite the work details mentioned previously. If you were Outside of Canada or your home country, then indicate your status in that country or territory').setItalic(true).setSpacingAfter(2);
    body.appendParagraph('PLEASE DO NOT LEAVE ANY GAPS IN TIME').setBold(true).setSpacingAfter(6);
    addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Activity (Job title /unemployed /Study / Travel / Retired)', 'City and Country', 'Status in Country', 'Name of company, employer/company, school, facility, as applicable'], data.spousePersonalHistoryRows || [], ['from', 'to', 'activity', 'cityCountry', 'statusInCountry', 'nameOfCompany']);
    addField(body, 'Have you ever been a member of any political, social, youth or student organization, trade unions and professional associations? (yes or no)', data.spouseGovtOrg);
    addField(body, 'Have you ever held a government position (such as civil servant, judge, police officer, employee in a security organization)? (yes or no)', data.spouseGovtPosition);
    addField(body, 'Have you ever served in military or paramilitary service for any country? (yes or no)', data.spouseMilitary);
    body.appendParagraph('').setSpacingAfter(8);
    addSectionHeader(body, '🏠 ADDRESS HISTORY');
    body.appendParagraph('List all the addresses where you have lived since your 18th Birthday or the Past 10 years, whichever is most recent. (Mention all address of inside and outside of Canada)').setItalic(true).setSpacingAfter(2);
    body.appendParagraph('NOTE: DO NOT LEAVE ANY GAPS AND MENTION FULL ADDRESS WITH POSTAL CODES.').setBold(true).setSpacingAfter(6);
    addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Street number and name', 'City or Town', 'Province, State or District', 'Postal code/ Zip code', 'Country or Territory'], data.spouseAddressHistoryRows || [], ['from', 'to', 'street', 'city', 'province', 'postalCode', 'country']);
    body.appendParagraph('').setSpacingAfter(8);
    addSectionHeader(body, '✈️ TRAVEL HISTORY');
    body.appendParagraph('List all trips you have taken outside of Canada or your home country in the last 10 years or since the age of 18. Include all trips: tourism, business, training etc.').setItalic(true).setSpacingAfter(2);
    body.appendParagraph('NOTE: LIST ALL THE COUNTRIES VISITED DURING YOUR TRANSIT/LAYOVER WHOSE STAMPS HAVE BEEN MARKED ON YOUR PASSPORT.').setBold(true).setSpacingAfter(6);
    addTableFromRows(body, ['From (YYYY-MM)', 'To (YYYY-MM)', 'Duration', 'Destination (City and Country)', 'Purpose of Visit', 'Provide details (if any)'], data.spouseTravelHistoryRowsPR || [], ['from', 'to', 'duration', 'destination', 'purpose', 'details']);
    body.appendParagraph('').setSpacingAfter(8);
  }

  addSectionHeader(body, '🗒️ Staff notes / observations: ( To be filled by our staff only - not applicant )');
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client\'s UCI number yourself.').setBold(true).setSpacingAfter(8);
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  [
    'Current Address: Kindly confirm with the client whether they reside in an apartment building or a house (it is not an apartment).',
    'Email Address: Please confirm with the client that the email address mentioned above is accurate.',
    'IRCC Refusals and Approvals: Kindly confirm with the client the details of previous application approvals and refusals.',
    'Have you confirmed that there is no month missing in the personal history section?',
    'Have you confirmed that the applicant did not miss any address? Physical location is very important to be declared, even if the address was not updated with the Bank, we still need to know about the physical address.',
    'Have you confirmed with the client about the names in native language (except Punjabi / Hindi)',
    'Have you confirmed about the travel history? Even one day of travel to any country has to be declared, if the immigration was done in that country.',
    'Have you confirmed with the applicant about his PCC. The PCC has to be done by specific departments only. Refer to the website for confirmations.',
    'Have you verified that there are no gaps in the personal history and address history sections?'
  ].forEach(function(n, i) {
    body.appendParagraph((i + 1) + '.  ' + n).setSpacingAfter(6);
  });
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}

// ── Visitor Visa: Personal, Sponsor, Spouse, Children, Education, Work,
// Family Information, Applications History (with biometrics). ──
function fillDocVisitorVisa(body, data) {
  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }
  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }
  function addSubHeader(body, title) {
    var p = body.appendParagraph(title);
    p.setBold(true).setItalic(true);
    p.setSpacingBefore(8).setSpacingAfter(4);
  }
  function addKeyValueTable(body, rows) {
    var tbl = body.appendTable(rows);
    tbl.setBorderWidth(1);
    for (var i = 0; i < rows.length; i++) {
      tbl.getCell(i, 0).setWidth(280).getChild(0).asParagraph().editAsText().setBold(true);
      tbl.getCell(i, 0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      tbl.getCell(i, 1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    }
  }
  function addTableFromRows(body, headers, rows, keys) {
    if (!rows || rows.length === 0) {
      body.appendTable([headers]).setBorderWidth(1);
      return;
    }
    body.appendTable([headers].concat(rows.map(function(r) {
      return keys.map(function(k) { return r[k] || '—'; });
    }))).setBorderWidth(1);
  }
  // The form always posts one blank row per repeatable block.
  function nonEmpty(rows) {
    if (!rows || !rows.length) return [];
    return rows.filter(function(r) {
      for (var k in r) {
        if (r.hasOwnProperty(k) && String(r[k] || '').trim() !== '') return true;
      }
      return false;
    });
  }

  var title = body.appendParagraph('INFORMATION SHEET (VISITOR VISA)');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);

  // ── Section 1 — Principal Applicant ──
  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var tableData = [
    ['Given Name', data.givenName || '—'],
    ['Last Name', data.lastName || '—'],
    ['Date of Birth (DD/MM/YYYY)', data.dob || '—'],
    ['Email ID', data.email || '—'],
    ['Phone Number', data.phone || '—'],
    ['Current Mailing Address (full with Postal code)', data.address || '—'],
    ['Marital Status', data.maritalStatus || '—'],
    ['Date of Marriage (DD/MM/YYYY)', data.dom || '—'],
    ['Do you have any children?', data.hasChildren || '—'],
    ['Native Language (Mother Tongue)', data.nativeLang || '—'],
    ['Status in Canada (current)', data.canadaStatus || '—'],
    ['Eye Color', data.eyeColor || '—'],
    ['Height (in cm)', data.heightCm || '—'],
    ['Passport No.', data.passport || '—'],
    ['UCI No.', data.uci || '—'],
    ['Have you ever used any other name? (yes/no)', data.otherNameUsed || '—']
  ];
  if (data.otherNameUsed === 'Yes') {
    tableData.push(['Other Name(s) Used', data.otherNames || '—']);
  }
  tableData.push(['Date of last entry to Canada', data.prLastEntryDate || '—']);
  tableData.push(['Place of last entry to Canada', data.prLastEntryPlace || '—']);
  tableData.push(['Have you previously been married or in a common-law relationship? (yes/no)', data.prPrevMarried || '—']);
  if (data.prPrevMarried === 'Yes') {
    tableData.push(['Previous Partner — Given Name', data.prPrevGiven || '—']);
    tableData.push(['Previous Partner — Last Name', data.prPrevLast || '—']);
    tableData.push(['Previous Partner — Date of Birth', data.prPrevDob || '—']);
    tableData.push(['Previous Partner — Type of Relationship', data.prPrevRelType || '—']);
    tableData.push(['Previous Partner — From', data.prPrevFrom || '—']);
    tableData.push(['Previous Partner — To', data.prPrevTo || '—']);
  }
  addKeyValueTable(body, tableData);
  body.appendParagraph('').setSpacingAfter(8);

  // ── Section 2 — Sponsor ──
  addSectionHeader(body, '🤝 SPONSOR INFORMATION');
  body.appendParagraph('Details of the person in Canada inviting or supporting this visit.').setItalic(true).setSpacingAfter(6);
  var sponsorRows = [
    ['Given Name', data.sponsorGiven || '—'],
    ['Last Name', data.sponsorLast || '—'],
    ['Relationship to Applicant', data.sponsorRelationship || '—'],
    ['Date of Birth (DD/MM/YYYY)', data.sponsorDob || '—'],
    ['Email ID', data.sponsorEmail || '—'],
    ['Phone Number', data.sponsorPhone || '—'],
    ['Current Mailing Address (full with Postal code)', data.sponsorAddress || '—'],
    ['Marital Status', data.sponsorMarital || '—'],
    ['Date of Marriage (DD/MM/YYYY)', data.sponsorDom || '—'],
    ['Do they have any children?', data.sponsorHasChildren || '—'],
    ['Native Language (Mother Tongue)', data.sponsorNativeLang || '—'],
    ['Status in Canada (current)', data.sponsorCanadaStatus || '—'],
    ['Passport No.', data.sponsorPassport || '—'],
    ['Current Employment', data.sponsorEmployment || '—'],
    ['Current Bank Balance', data.sponsorBankBalance || '—'],
    ['Have they ever used any other name? (yes/no)', data.sponsorOtherNameUsed || '—']
  ];
  if (data.sponsorOtherNameUsed === 'Yes') {
    sponsorRows.push(['Other Name(s) Used', data.sponsorOtherNames || '—']);
  }
  sponsorRows.push(['Date of last entry to Canada', data.sponsorLastEntryDate || '—']);
  sponsorRows.push(['Place of last entry to Canada', data.sponsorLastEntryPlace || '—']);
  sponsorRows.push(['Have they previously been married or in a common-law relationship? (yes/no)', data.sponsorPrevMarried || '—']);
  if (data.sponsorPrevMarried === 'Yes') {
    sponsorRows.push(['Previous Partner — Given Name', data.sponsorPrevGiven || '—']);
    sponsorRows.push(['Previous Partner — Last Name', data.sponsorPrevLast || '—']);
    sponsorRows.push(['Previous Partner — Date of Birth', data.sponsorPrevDob || '—']);
    sponsorRows.push(['Previous Partner — Type of Relationship', data.sponsorPrevRelType || '—']);
    sponsorRows.push(['Previous Partner — From', data.sponsorPrevFrom || '—']);
    sponsorRows.push(['Previous Partner — To', data.sponsorPrevTo || '—']);
  }
  addKeyValueTable(body, sponsorRows);
  body.appendParagraph('').setSpacingAfter(8);

  // ── Section 3 — Spouse ──
  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    addField(body, 'Given Name', data.spouseGiven);
    addField(body, 'Last Name', data.spouseLast);
    addField(body, 'Date of Birth (DD/MM/YYYY)', data.spouseDob);
    addField(body, 'Email ID', data.spouseEmail);
    addField(body, 'Phone Number', data.spousePhone);
    addField(body, 'Current Address', data.spouseAddr);
    addField(body, 'Status in Canada (if in Canada)', data.spouseCanadaStatus);
    addField(body, 'Eye Color', data.spouseEyeColor);
    addField(body, 'Height (in cm)', data.spouseHeightCm);
    addField(body, 'Have you been married or in a common-law relationship before your current marriage?', data.prevRel);
    if (data.prevRel === 'Yes') {
      addSubHeader(body, 'If yes, kindly provide the following details:');
      addField(body, 'Given Name', data.prevGiven);
      addField(body, 'Last Name', data.prevLast);
      addField(body, 'Date of Birth (DD/MM/YYYY)', data.prevDob);
      addField(body, 'Type of Relationship', data.prevType);
      addField(body, 'From', data.prevFrom);
      addField(body, 'To', data.prevTo);
    }
    body.appendParagraph('').setSpacingAfter(8);
  }

  // ── Section 4 — Children ──
  if (showChildrenSection_(data)) {
    addSectionHeader(body, '🔹 CHILDREN INFO (If Any)');
    addField(body, '1', data.child1);
    addField(body, '2', data.child2);
    addField(body, '3', data.child3);
    if (data.childExtra) addField(body, 'Additional', data.childExtra);
  }

  // ── Section 5 — Education ──
  addSectionHeader(body, '🎓 EDUCATION');
  body.appendParagraph('Grade 12 and all post-secondary programs — completed or not, inside or outside of Canada.').setItalic(true).setSpacingAfter(6);
  body.appendParagraph('Grade 12 (High School)').setBold(true);
  addField(body, 'From (MM/YYYY)', data.g12from);
  addField(body, 'To (MM/YYYY)', data.g12to);
  addField(body, 'Program / Stream', data.g12prog);
  addField(body, 'School Name', data.g12inst);
  addField(body, 'City', data.g12city);
  body.appendParagraph('').setSpacingAfter(4);

  var eduEntries = (data.education || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Entry \d+:\s*/, '').trim();
    return c.replace(/to\s*\|\s*\|\s*\|/, '').replace(/\|/g, '').trim() !== '';
  });
  if (eduEntries.length > 0) {
    addSubHeader(body, 'Post-Secondary Education');
    eduEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Entry \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      body.appendParagraph('Education ' + (idx + 1)).setBold(true);
      addField(body, 'From (MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Program Name', parts[1] || '—');
      addField(body, 'Institute Name', parts[2] || '—');
      addField(body, 'Campus City', parts[3] || '—');
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // ── Section 6 — Work History ──
  addSectionHeader(body, '💼 WORK HISTORY (All works done on SIN in Canada and all works done outside of Canada)');
  var workEntries = (data.work || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Job \d+:\s*/, '').trim();
    return c.replace(/\|/g, '').trim() !== '';
  });
  if (workEntries.length === 0) { body.appendParagraph('—'); }
  else {
    workEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Job \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      var isNewFormat = parts.length >= 6;
      body.appendParagraph((idx === 0 ? '1️⃣ Current employment details:' : (idx + 1) + '️⃣')).setBold(true);
      addField(body, 'From (DD/MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (DD/MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Job Title', parts[1] || '—');
      if (isNewFormat) {
        addField(body, 'Full-time or Part-time', parts[2] || '—');
        addField(body, 'Hours per week', parts[3] || '—');
        addField(body, 'Employer / Company Name', parts[4] || '—');
        addField(body, 'Full Employer / Work Location Address', parts[5] || '—');
      } else {
        addField(body, 'Employer / Company Name', parts[2] || '—');
        addField(body, 'Full Employer / Work Location Address', parts[3] || '—');
      }
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // ── Section 7 — Family Information ──
  addSectionHeader(body, '👪 FAMILY INFORMATION');
  body.appendParagraph('Family members (parents, siblings, spouse, children). If deceased, date of death in address field.').setItalic(true).setSpacingAfter(6);
  addTableFromRows(body,
    ['Full Name', 'Date of Birth', 'Place of Birth', 'Marital Status', 'Relationship', 'Email', 'Occupation', 'Current Address'],
    nonEmpty(data.familyMembers),
    ['fullName', 'dob', 'placeOfBirth', 'maritalStatus', 'relationship', 'email', 'occupation', 'currentAddress']
  );
  body.appendParagraph('').setSpacingAfter(8);

  // ── Section 8 — Applications History and Background ──
  addSectionHeader(body, '🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  body.appendParagraph('List all the applications ever applied to IRCC: Study permit, work permit, extensions of all kinds, TRV, visitor.').setItalic(true).setSpacingAfter(6);
  var appRows = nonEmpty(data.appHistoryRows);
  var appHeaderRow = ['Type of application', 'Result (Approved/denied)', 'Date of result', 'Destination in Canada'];
  if (appRows.length === 0) {
    body.appendTable([appHeaderRow]).setBorderWidth(1);
  } else {
    body.appendTable([appHeaderRow].concat(appRows.map(function(r) {
      return [r.type || '—', r.result || '—', r.date || '—', r.destination || '—'];
    }))).setBorderWidth(1);
  }
  body.appendParagraph('').setSpacingAfter(8);
  addField(body, 'Have you ever been refused by IRCC or (US, AUS, NZ or any other country)?', data.refused);
  if (data.refused === 'Yes') {
    addField(body, 'Country', data.refCountry);
    addField(body, 'Result Date (MM/YYYY)', data.refDate);
    addField(body, 'Type of Application (Visit / Work / PR)', data.refType);
  }
  addField(body, 'Have you ever committed, been arrested for or been charged with or convicted of any criminal offense in any country?', data.criminalRecord);
  if (data.criminalRecord === 'Yes') {
    addField(body, 'Details', data.criminalInfo);
  }
  addField(body, 'Have you ever submitted biometrics?', data.biometricsGiven);
  if (data.biometricsGiven === 'Yes') {
    addField(body, 'Date biometrics were given', data.biometricsDate);
  }

  // ── Staff Notes ──
  addSectionHeader(body, '🗒️ Staff Notes / Observations');
  body.appendParagraph('To be filled by our staff only — not applicant').setItalic(true).setSpacingAfter(8);
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client\'s UCI number yourself.').setBold(true).setSpacingAfter(8);
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  [
    'Current Address: Kindly confirm with the client whether they reside in an apartment building or a house.',
    'Email Address: Please confirm with the client that the email address mentioned above is accurate.',
    'Sponsor: Confirm the sponsor\'s status in Canada and that the relationship to the applicant is documented.',
    'Sponsor Funds: Confirm the bank balance is supported by statements on file.',
    'IRCC Refusals and Approvals: Kindly confirm with the client the details of previous application approvals and refusals.',
    'Biometrics: If biometrics were given, confirm they are still valid (10 years) and note the date.',
    'Policy and Procedure: Did you read the policy and any update related to the type of application?',
    'Documents Checklist: Did you make sure that all the documents in the checklist are uploaded and marked with ✅? If anything is missing, please explain the reason.',
    'Passport Expiry: Have you checked if the passport expiry date is more than 3 years? If not, kindly highlight in this section.'
  ].forEach(function(n, i) {
    body.appendParagraph((i + 1) + '.  ' + n).setSpacingAfter(6);
  });
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}

// ── PR (Permanent Residence): new application type added via form ──
function fillDocPR(body, data) {
  function addSectionHeader(body, title) {
    body.appendParagraph('').setSpacingAfter(2);
    var h = body.appendParagraph(title);
    h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    h.setSpacingBefore(14);
    h.setSpacingAfter(6);
  }
  function addField(body, label, value) {
    var p = body.appendParagraph('');
    p.setSpacingAfter(5).setSpacingBefore(0);
    p.appendText(label + ':  ').setBold(true);
    p.appendText(value || '—');
  }
  function addSubHeader(body, title) {
    var p = body.appendParagraph(title);
    p.setBold(true).setItalic(true);
    p.setSpacingBefore(8).setSpacingAfter(4);
  }
  function addTableFromRows(body, headers, rows, keys) {
    if (!rows || rows.length === 0) {
      var tbl = body.appendTable([headers]);
      tbl.setBorderWidth(1);
      return;
    }
    var tableData = [headers].concat(rows.map(function(r) { return keys.map(function(k) { return r[k] || '—'; }); }));
    var tbl = body.appendTable(tableData);
    tbl.setBorderWidth(1);
  }

  // Title
  var title = body.appendParagraph('INFORMATION SHEET (PR)');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setSpacingAfter(4);

  // Section 1 — Principal Applicant
  addSectionHeader(body, '🔹 CLIENT INFORMATION (PRINCIPAL APPLICANT)');
  var tableData = [
    ['Given Name', data.givenName||'—'],
    ['Last Name', data.lastName||'—'],
    ['Date of Birth (DD/MM/YYYY)', data.dob||'—'],
    ['Email ID', data.email||'—'],
    ['Phone Number', data.phone||'—'],
    ['Current Residential Address (full with Postal code)', data.address||'—'],
    ['Marital Status', data.maritalStatus||'—'],
    ['Date of Marriage (DD/MM/YYYY)', data.dom||'—'],
    ['Native Language (Mother Tongue)', data.nativeLang||'—'],
    ['Status in Canada (current)', data.canadaStatus||'—'],
    ['Eye Color', data.eyeColor||'—'],
    ['Height (in cm)', data.heightCm||'—'],
    ['Passport No.', data.passport||'—'],
    ['UCI No.', data.uci||'—'],
    ['Have you ever used any other name? (yes/no)', data.otherNameUsed||'—']
  ];
  if (data.otherNameUsed === 'Yes') {
    tableData.push(['Other Name(s) Used', data.otherNames||'—']);
  }
  tableData.push(['Date of last entry to Canada', data.prLastEntryDate||'—']);
  tableData.push(['Place of last entry to Canada', data.prLastEntryPlace||'—']);
  tableData.push(['Have you previously been married or in a common-law relationship? (yes/no)', data.prPrevMarried||'—']);
  if (data.prPrevMarried === 'Yes') {
    tableData.push(['Previous Partner — Given Name', data.prPrevGiven||'—']);
    tableData.push(['Previous Partner — Last Name', data.prPrevLast||'—']);
    tableData.push(['Previous Partner — Date of Birth', data.prPrevDob||'—']);
    tableData.push(['Previous Partner — Type of Relationship', data.prPrevRelType||'—']);
    tableData.push(['Previous Partner — From', data.prPrevFrom||'—']);
    tableData.push(['Previous Partner — To', data.prPrevTo||'—']);
  }
  var table = body.appendTable(tableData);
  table.setBorderWidth(1);
  for (var r = 0; r < tableData.length; r++) {
    table.getCell(r, 0).setWidth(280).getChild(0).asParagraph().editAsText().setBold(true);
    table.getCell(r, 0).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
    table.getCell(r, 1).setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
  }
  body.appendParagraph('').setSpacingAfter(8);

  // Section 2 — Spouse
  if (showSpouseSection_(data)) {
    addSectionHeader(body, '🔹 SPOUSE / COMMON-LAW PARTNER INFO (If Any)');
    addField(body, 'Given Name', data.spouseGiven);
    addField(body, 'Last Name', data.spouseLast);
    addField(body, 'Date of Birth (DD/MM/YYYY)', data.spouseDob);
    addField(body, 'Email ID', data.spouseEmail);
    addField(body, 'Phone Number', data.spousePhone);
    addField(body, 'Current Address', data.spouseAddr);
    addField(body, 'Status in Canada (if in Canada)', data.spouseCanadaStatus);
    addField(body, 'Eye Color', data.spouseEyeColor);
    addField(body, 'Height (in cm)', data.spouseHeightCm);
    addField(body, 'Have you been married or in a common-law relationship before your current marriage?', data.prevRel);
    if (data.prevRel === 'Yes') {
      addSubHeader(body, 'If yes, kindly provide the following details:');
      addField(body, 'Given Name', data.prevGiven);
      addField(body, 'Last Name', data.prevLast);
      addField(body, 'Date of Birth (DD/MM/YYYY)', data.prevDob);
      addField(body, 'Type of Relationship', data.prevType);
      addField(body, 'From', data.prevFrom);
      addField(body, 'To', data.prevTo);
    }
    body.appendParagraph('').setSpacingAfter(8);
  }

  // Section 3 — Children
  if (showChildrenSection_(data)) {
    addSectionHeader(body, '🔹 CHILDREN INFO (If Any)');
    addField(body, '1', data.child1);
    addField(body, '2', data.child2);
    addField(body, '3', data.child3);
    if (data.childExtra) addField(body, 'Additional', data.childExtra);
  }

  // Section 4 — Travel (basic)
  addSectionHeader(body, '✈️ TRAVEL HISTORY');
  addField(body, 'First Entry to Canada (Date)', data.firstEntryDate);
  addField(body, 'First Entry to Canada (Airport)', data.firstEntryPort);
  addField(body, 'Recent Entry to Canada (Date)', data.recentEntryDate);
  addField(body, 'Recent Entry to Canada (Place)', data.recentEntryPort);

  // Section 5 — Education
  addSectionHeader(body, '🎓 EDUCATION');
  body.appendParagraph('Grade 12 and all post-secondary programs — completed or not, inside or outside of Canada.').setItalic(true).setSpacingAfter(6);

  // Grade 12
  body.appendParagraph('Grade 12 (High School)').setBold(true);
  addField(body, 'From (MM/YYYY)', data.g12from);
  addField(body, 'To (MM/YYYY)', data.g12to);
  addField(body, 'Program / Stream', data.g12prog);
  addField(body, 'School Name', data.g12inst);
  addField(body, 'City', data.g12city);
  if (data.g12level) addField(body, 'Level of Education', data.g12level);
  if (data.g12field) addField(body, 'Field of Study', data.g12field);
  body.appendParagraph('').setSpacingAfter(4);

  // Post-secondary — parse from pipe-delimited string + educationExtras array
  var eduEntries = (data.education || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Entry \d+:\s*/, '').trim();
    return c.replace(/to\s*\|\s*\|\s*\|/, '').replace(/\|/g, '').trim() !== '';
  });
  var eduExtras = data.educationExtras || [];
  if (eduEntries.length > 0) {
    addSubHeader(body, 'Post-Secondary Education');
    eduEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Entry \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      var extra = eduExtras[idx] || {};
      body.appendParagraph('Education ' + (idx + 1)).setBold(true);
      addField(body, 'From (MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Program Name', parts[1] || '—');
      addField(body, 'Institute Name', parts[2] || '—');
      addField(body, 'Campus City', parts[3] || '—');
      if (extra.level) addField(body, 'Level of Education', extra.level);
      if (extra.field) addField(body, 'Field of Study', extra.field);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 6 — Work History (with contact person extras for PR)
  addSectionHeader(body, '💼 WORK HISTORY');
  var workEntries = (data.work || '').split('\n').filter(function(e) {
    if (!e.trim()) return false;
    var c = e.replace(/^Job \d+:\s*/, '').trim();
    return c.replace(/\|/g, '').trim() !== '';
  });
  var workExtras = data.workContactExtras || [];
  if (workEntries.length === 0) { body.appendParagraph('—'); }
  else {
    workEntries.forEach(function(entry, idx) {
      var parts = entry.replace(/^Job \d+:\s*/, '').split(' | ');
      var dates = (parts[0] || '').split(' to ');
      var extra = workExtras[idx] || {};
      body.appendParagraph((idx === 0 ? '1️⃣ Current employment details:' : (idx + 1) + '️⃣')).setBold(true);
      addField(body, 'From (DD/MM/YYYY)', dates[0] ? dates[0].trim() : '—');
      addField(body, 'To (DD/MM/YYYY)', dates[1] ? dates[1].trim() : '—');
      addField(body, 'Job Title', parts[1] || '—');
      addField(body, 'Full-time or Part-time', parts[2] || '—');
      addField(body, 'Hours per week', parts[3] || '—');
      addField(body, 'Employer / Company Name', parts[4] || '—');
      addField(body, 'Full Work Location Address', parts[5] || '—');
      if (extra.contactName) addField(body, 'Contact Person Name', extra.contactName);
      if (extra.contactPhone) addField(body, 'Contact Person Phone', extra.contactPhone);
      if (extra.contactEmail) addField(body, 'Contact Email', extra.contactEmail);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Section 7 — Financial
  addSectionHeader(body, '💰 Financial Questions');
  addField(body, 'How much do you have in savings or checking accounts?', data.savings);
  addField(body, 'If you do not have a job, who and how are you paying for the expenses?', data.expenses);

  // Section 8 — Application History
  addSectionHeader(body, '🗂️ APPLICATIONS HISTORY AND BACKGROUND');
  body.appendParagraph('List all applications submitted to IRCC.').setItalic(true).setSpacingAfter(6);
  var appRows = data.appHistoryRows || [];
  var appHeaderRow = ['Type of application', 'Result (Approved/denied)', 'Date of result', 'Destination in Canada', 'Reason for refusal (if known)'];
  if (appRows.length === 0) {
    body.appendTable([appHeaderRow]).setBorderWidth(1);
  } else {
    var appTableData = [appHeaderRow].concat(appRows.map(function(r) {
      return [r.type || '—', r.result || '—', r.date || '—', r.destination || '—', r.reason || '—'];
    }));
    body.appendTable(appTableData).setBorderWidth(1);
  }
  body.appendParagraph('').setSpacingAfter(8);
  addField(body, 'Have you ever been refused by IRCC or (US, AUS, NZ or any other country)?', data.refused);
  if (data.refused === 'Yes') {
    addField(body, 'Country', data.refCountry);
    addField(body, 'Result Date (MM/YYYY)', data.refDate);
    addField(body, 'Type of Application (Visit / Work / PR)', data.refType);
  }
  addField(body, 'Have you ever committed, been arrested for or been charged with or convicted of any criminal offense in any country?', data.criminalRecord);
  if (data.criminalRecord === 'Yes') {
    addField(body, 'Details', data.criminalInfo);
  }

  // Section 9 — Family Information
  addSectionHeader(body, '👪 Family Information');
  body.appendParagraph('Family members (parents, siblings, spouse, children). If deceased, date of death in address field.').setItalic(true).setSpacingAfter(6);
  var familyHeaders = ['Full Name', 'Date of Birth', 'Place of Birth', 'Marital Status', 'Relationship', 'Email', 'Occupation', 'Current Address'];
  var familyKeys = ['fullName', 'dob', 'placeOfBirth', 'maritalStatus', 'relationship', 'email', 'occupation', 'currentAddress'];
  addTableFromRows(body, familyHeaders, data.familyMembers || [], familyKeys);
  body.appendParagraph('').setSpacingAfter(8);

  // Section 10 — Personal History (PR)
  addSectionHeader(body, '📝 PERSONAL HISTORY');
  body.appendParagraph('Details since age 18 or past 10 years. Under activity, write occupation/job title or what you were doing (unemployed, studying, travelling, retired, etc.)').setItalic(true).setSpacingAfter(2);
  body.appendParagraph('PLEASE DO NOT LEAVE ANY GAPS IN TIME').setBold(true).setSpacingAfter(6);
  addTableFromRows(body,
    ['From (YYYY-MM)', 'To (YYYY-MM)', 'Activity', 'City and Country', 'Status in Country', 'Name of Company/Employer/School'],
    data.personalHistory || [],
    ['from', 'to', 'activity', 'city', 'status', 'company']
  );
  addField(body, 'Member of any political/social/youth/student organization, trade unions or professional associations?', data.bgPolitical);
  addField(body, 'Ever held a government position (civil servant, judge, police officer, security organization)?', data.bgGovt);
  addField(body, 'Ever served in military or paramilitary service?', data.bgMilitary);
  body.appendParagraph('').setSpacingAfter(8);

  // Section 11 — Address History (PR)
  addSectionHeader(body, '🏠 ADDRESS HISTORY');
  body.appendParagraph('All addresses since 18th birthday or past 10 years. DO NOT LEAVE ANY GAPS. MENTION FULL ADDRESS WITH POSTAL CODES.').setBold(true).setSpacingAfter(6);
  addTableFromRows(body,
    ['From (YYYY-MM-DD)', 'To (YYYY-MM-DD)', 'Street Number and Name', 'City or Town', 'Province/State/District', 'Postal/Zip Code', 'Country'],
    data.addressHistory || [],
    ['from', 'to', 'street', 'city', 'province', 'postal', 'country']
  );
  body.appendParagraph('').setSpacingAfter(8);

  // Section 12 — Travel History (Detailed, PR)
  addSectionHeader(body, '✈️ TRAVEL HISTORY (Detailed)');
  body.appendParagraph('All trips outside Canada or home country in last 10 years or since age 18. Include transit/layover countries with passport stamps.').setItalic(true).setSpacingAfter(6);
  addTableFromRows(body,
    ['From (YYYY-MM)', 'To (YYYY-MM)', 'Duration', 'Destination (City and Country)', 'Purpose of Visit', 'Details'],
    data.travelHistory || [],
    ['from', 'to', 'duration', 'destination', 'purpose', 'details']
  );
  body.appendParagraph('').setSpacingAfter(8);

  // ── SPOUSE SECTIONS (only if any spouse section has data) ──
  var hasSpouseSections = (data.spouseFamily && data.spouseFamily.length > 0) ||
    (data.spouseAppHistory && data.spouseAppHistory.length > 0) ||
    (data.spouseEducation && data.spouseEducation.length > 0) ||
    (data.spouseWork && data.spouseWork.length > 0) ||
    (data.spousePersonalHistory && data.spousePersonalHistory.length > 0) ||
    (data.spouseAddrHistory && data.spouseAddrHistory.length > 0) ||
    (data.spouseTravelHistory && data.spouseTravelHistory.length > 0);

  if (hasSpouseSections) {
    body.appendParagraph('IF APPLYING FOR YOUR SPOUSE ALSO, THEN PROVIDE THE FOLLOWING INFORMATION FOR SPOUSE:').setBold(true).setSpacingBefore(16).setSpacingAfter(8);

    // Spouse Family
    if (data.spouseFamily && data.spouseFamily.length > 0) {
      addSectionHeader(body, '👫 Spouse — Family Information');
      body.appendParagraph('Family members of the spouse (parents, children, siblings). If deceased, provide date of death in address field.').setItalic(true).setSpacingAfter(6);
      var spFamHeaders = ['Full Name', 'Date of Birth', 'Place of Birth', 'Marital Status', 'Relationship', 'Email', 'Current Address'];
      addTableFromRows(body, spFamHeaders, data.spouseFamily, ['fullName', 'dob', 'placeOfBirth', 'maritalStatus', 'relationship', 'email', 'currentAddress']);
      body.appendParagraph('').setSpacingAfter(8);
    }

    // Spouse App History
    if (data.spouseAppHistory && data.spouseAppHistory.length > 0) {
      addSectionHeader(body, '🗂️ Spouse — Application History');
      addTableFromRows(body,
        ['Type of application', 'Result', 'Date of result', 'Destination in Canada', 'Reason for refusal'],
        data.spouseAppHistory,
        ['type', 'result', 'date', 'destination', 'reason']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }

    // Spouse Education
    if (data.spouseEducation && data.spouseEducation.length > 0) {
      addSectionHeader(body, '🎓 Spouse — Education History');
      addTableFromRows(body,
        ['From (YYYY-MM)', 'To (YYYY-MM)', 'Institution', 'City and Country', 'Level of Education', 'Field of Study'],
        data.spouseEducation,
        ['from', 'to', 'institution', 'city', 'level', 'field']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }

    // Spouse Work
    if (data.spouseWork && data.spouseWork.length > 0) {
      addSectionHeader(body, '💼 Spouse — Work History');
      data.spouseWork.forEach(function(w, idx) {
        body.appendParagraph('Spouse Job ' + (idx + 1)).setBold(true);
        addField(body, 'From', w.from);
        addField(body, 'To', w.to);
        addField(body, 'Job Title', w.title);
        addField(body, 'Employer / Company Name', w.employer);
        addField(body, 'Full Work Location Address', w.address);
        if (w.contactName) addField(body, 'Contact Person Name', w.contactName);
        if (w.contactPhone) addField(body, 'Contact Person Phone', w.contactPhone);
        if (w.contactEmail) addField(body, 'Contact Email', w.contactEmail);
        body.appendParagraph('').setSpacingAfter(4);
      });
    }

    // Spouse Personal History
    if (data.spousePersonalHistory && data.spousePersonalHistory.length > 0) {
      addSectionHeader(body, '📝 Spouse — Personal History');
      body.appendParagraph('PLEASE DO NOT LEAVE ANY GAPS IN TIME').setBold(true).setSpacingAfter(6);
      addTableFromRows(body,
        ['From (YYYY-MM)', 'To (YYYY-MM)', 'Activity', 'City and Country', 'Status in Country', 'Company/Employer/School'],
        data.spousePersonalHistory,
        ['from', 'to', 'activity', 'city', 'status', 'company']
      );
      addField(body, 'Member of any political/social/youth/student organization?', data.spBgPolitical);
      addField(body, 'Ever held a government position?', data.spBgGovt);
      addField(body, 'Ever served in military or paramilitary service?', data.spBgMilitary);
      body.appendParagraph('').setSpacingAfter(8);
    }

    // Spouse Address History
    if (data.spouseAddrHistory && data.spouseAddrHistory.length > 0) {
      addSectionHeader(body, '🏠 Spouse — Address History');
      body.appendParagraph('DO NOT LEAVE ANY GAPS. MENTION FULL ADDRESS WITH POSTAL CODES.').setBold(true).setSpacingAfter(6);
      addTableFromRows(body,
        ['From (YYYY-MM-DD)', 'To (YYYY-MM-DD)', 'Street', 'City or Town', 'Province/State', 'Postal/Zip Code', 'Country'],
        data.spouseAddrHistory,
        ['from', 'to', 'street', 'city', 'province', 'postal', 'country']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }

    // Spouse Travel History
    if (data.spouseTravelHistory && data.spouseTravelHistory.length > 0) {
      addSectionHeader(body, '✈️ Spouse — Travel History');
      addTableFromRows(body,
        ['From (YYYY-MM)', 'To (YYYY-MM)', 'Duration', 'Destination', 'Purpose', 'Details'],
        data.spouseTravelHistory,
        ['from', 'to', 'duration', 'destination', 'purpose', 'details']
      );
      body.appendParagraph('').setSpacingAfter(8);
    }
  }

  // Relatives in Canada (PR)
  if (data.prRelCan === 'Yes' && data.prRelativesCanada && data.prRelativesCanada.length > 0) {
    addSectionHeader(body, '👪 Relatives in Canada');
    data.prRelativesCanada.forEach(function(rel, idx) {
      body.appendParagraph('Relative ' + (idx + 1)).setBold(true);
      addField(body, 'Name', rel.name);
      addField(body, 'Date of Birth', rel.dob);
      addField(body, 'Email ID', rel.email);
      addField(body, 'Phone Number', rel.phone);
      addField(body, 'Relationship to Applicant', rel.relationship);
      addField(body, 'Current Address', rel.address);
      addField(body, 'Status in Canada', rel.status);
      addField(body, 'Date they became PR (if PR or Citizen)', rel.prDate);
      addField(body, 'Date they moved to MB (if PR or Citizen)', rel.mbDate);
      body.appendParagraph('').setSpacingAfter(4);
    });
  }

  // Staff Notes
  addSectionHeader(body, '🗒️ Staff Notes / Observations (Staff Use Only)');
  body.appendParagraph('📌 UCI Number: The UCI number is not collected on this form. Please look up and add the client\'s UCI number yourself.').setBold(true).setSpacingAfter(8);
  body.appendParagraph('Information to Be Confirmed (For Staff Use Only):').setBold(true);
  var staffChecks = [
    ['Current Address confirmed (apartment or house)?', data.staffAddr],
    ['Email Address confirmed as accurate?', data.staffEmail],
    ['IRCC Refusals and Approvals confirmed?', data.staffIRCC],
    ['Personal History — no month missing?', data.staffPH],
    ['Address History — no address missed?', data.staffAH],
    ['Native Language names confirmed (except Punjabi/Hindi)?', data.staffLang],
    ['Travel History — even one-day travel declared?', data.staffTravel],
    ['PCC confirmed from specific departments?', data.staffPCC],
    ['No gaps in personal history and address history?', data.staffGaps]
  ];
  staffChecks.forEach(function(item, i) {
    body.appendParagraph((i + 1) + '.  ' + item[0] + '  ' + (item[1] || '—')).setSpacingAfter(6);
  });
  if (data.staffMethod) addField(body, 'Method used to confirm information', data.staffMethod);
  if (data.staffNotes) addField(body, 'Additional Staff Notes', data.staffNotes);
  body.appendParagraph('').setSpacingAfter(4);
  body.appendParagraph('Note: Please ensure that these details are confirmed with the client either in a group or through a personal call. Additionally, kindly specify the method used to confirm the information.').setItalic(true);
}