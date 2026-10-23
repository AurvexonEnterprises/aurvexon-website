/** Aurvexon Talent shared registration endpoint.
 * Use the same spreadsheet as registration.aurvexon.in.
 * Deploy as Web app: Execute as Me; Who has access: Anyone.
 */
const SPREADSHEET_ID = '1ppKLRyXUXbV3k67Jf0u2tJbhrXyvqd8a2Y_8ASYefE4';
const CANDIDATE_SHEET = 'Candidate Registrations';
const EMPLOYER_SHEET = 'Employer Enquiries';

const CANDIDATE_HEADERS = [
  'Server timestamp', 'Full name', 'WhatsApp number', 'Email',
  'Preferred role', 'Preferred location', 'Expected monthly salary',
  'Terms accepted', 'Privacy consent', 'Terms version', 'Client timestamp', 'Review status',
  'Phone number', 'Current location', 'Highest qualification', 'Work experience',
  'Preferred job location', 'Preferred work type', 'CV / Resume link', 'Additional details'
];
const EMPLOYER_HEADERS = [
  'Server timestamp', 'Company', 'Contact name', 'Email', 'Phone',
  'Service required', 'Number of openings', 'Role details', 'Status'
];

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ok:true, service:'Aurvexon Talent shared registration endpoint'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let legacyCandidatePage = false;
  let result;
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    let data = params;
    let type = String(params.formType || '').toLowerCase();

    // The separate registration.aurvexon.in page sends its candidate fields as JSON in `payload`.
    if (params.payload) {
      data = JSON.parse(params.payload);
      type = 'candidate';
      legacyCandidatePage = true;
    }
    if (type !== 'candidate' && type !== 'employer') throw new Error('Invalid form type.');

    const required = type === 'candidate'
      ? (legacyCandidatePage
          ? ['fullName','whatsapp','preferredRole','preferredLocation','expectedSalary']
          : ['fullName','email','phone','location','qualification','experience','preferredRole','preferredJobLocation'])
      : ['company','contactName','email','phone','service','details'];
    required.forEach(key => {
      if (!String(data[key] || '').trim()) throw new Error('Please complete the required field: ' + key);
    });
    if (legacyCandidatePage && (data.termsConsent !== true || data.privacyConsent !== true)) {
      throw new Error('Required consent was not provided.');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = type === 'candidate' ? CANDIDATE_SHEET : EMPLOYER_SHEET;
    const baseHeaders = type === 'candidate' ? CANDIDATE_HEADERS : EMPLOYER_HEADERS;
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(baseHeaders);
      sheet.setFrozenRows(1);
    } else {
      // Keep existing columns and add any new fields at the end, without overwriting old data.
      const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      const missing = baseHeaders.filter(h => !existing.includes(h));
      if (missing.length) sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }

    const safe = value => {
      const s = String(value == null ? '' : value).trim();
      return /^[=+@\-]/.test(s) ? "'" + s : s;
    };
    const rowData = type === 'candidate' ? {
      'Server timestamp': new Date(),
      'Full name': data.fullName,
      'WhatsApp number': data.whatsapp || data.phone,
      'Email': data.email,
      'Preferred role': data.preferredRole,
      'Preferred location': data.preferredLocation || data.preferredJobLocation,
      'Expected monthly salary': data.expectedSalary,
      'Terms accepted': legacyCandidatePage ? 'Yes' : (data.termsConsent ? 'Yes' : ''),
      'Privacy consent': legacyCandidatePage ? 'Yes' : (data.privacyConsent ? 'Yes' : ''),
      'Terms version': data.termsVersion || '',
      'Client timestamp': data.submittedAtClient || '',
      'Review status': 'Pending review',
      'Phone number': data.phone || data.whatsapp,
      'Current location': data.location,
      'Highest qualification': data.qualification,
      'Work experience': data.experience,
      'Preferred job location': data.preferredJobLocation,
      'Preferred work type': data.workType,
      'CV / Resume link': data.cvLink,
      'Additional details': data.notes
    } : {
      'Server timestamp': new Date(),
      'Company': data.company,
      'Contact name': data.contactName,
      'Email': data.email,
      'Phone': data.phone,
      'Service required': data.service,
      'Number of openings': data.openings,
      'Role details': data.details,
      'Status': 'Pending review'
    };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const row = headers.map(header => Object.prototype.hasOwnProperty.call(rowData, header)
      ? (header === 'Server timestamp' ? rowData[header] : safe(rowData[header]))
      : '');
    sheet.appendRow(row);
    result = {marker:'AURVEXON_REGISTRATION_RESULT', ok:true, formType:type};
  } catch (err) {
    result = {marker:'AURVEXON_REGISTRATION_RESULT', ok:false, message:String(err && err.message || err)};
  }

  // Preserve the postMessage response expected by the separate registration.aurvexon.in page.
  if (legacyCandidatePage) {
    return HtmlService.createHtmlOutput('<!doctype html><html><head><meta charset="utf-8"></head><body><script>window.top.postMessage(' + JSON.stringify(result) + ', "*");</script></body></html>')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
