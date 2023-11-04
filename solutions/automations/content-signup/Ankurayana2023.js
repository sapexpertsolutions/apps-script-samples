const sheetName = "FormResponses1"; // replace this sheet name with yours
const uidHeader = "Registration No"; // column name of UID in response sheet
const sLogHeader = "Email Sent Log"; //column name for email log
const uidPrefix = "ANK2023-"; // UID prefeix
const uidDigitLength = 4; // digit length of UID

function _onFormSubmit(e) {
  const { rowStart } = e.range
  const url = SpreadsheetApp.getActive().getFormUrl();
  const form = FormApp.openByUrl(url);
  const response = form.getResponses().pop();
  const url = form.getPublishedUrl();
  const editUrl = response.getEditResponseUrl()
  const responseId = response.getId();

  // for sequential id
  const uid = createUid(responseId);
  // for random id
  //const uid = createRandomUid(responseId);
  writeUidToSheet(uid, rowStart);

  const email = response.getRespondentEmail();
  if (email) {

    const AnkVersion = "Ankurayan 2023"; // Ankurayan Verion
    const AnkContactNo = "9778249509, 9999"; //Contact Numbers
    const AnkPunchLine = "ଅଙ୍କୁରାୟନ ହସର ଫଗୁଣ - ମାଟିରେ ମଗନ"; //Anlrayan Punch Line
    const BandhuContactNo = "9778249509, 9778249509"; //Bandhu office Contact No
    const greet = "Greetings from Bandhu family!!!";
    const title = "Thanks for your registration in ";
    const contmob = "For any enquiry and feedback, you can contact on :  "

    const subject = "Bandhu : Registration Number ";

    const htmlBody = "<p><table style='width: 600px; margin: 10px auto; background: #dfe4ea;'><tr>" +
      "<td style='padding: 0px 20px;'> <font face='verdana'>" + greet + "<br>" + title + AnkVersion +
      "<p>Please refer your Unique Registration No</p>" +
      "<div style='padding: 20px;border-radius: 20px; background: linear-gradient(60deg, #609fd6, #1aafbc);'>" +

      "<div style='font-size: 20px; color: #E06666; text-align: center;'><b>" + uid + "</b></div>" +

      "</div>" + "<br>" + contmob + AnkContactNo +

      "<p>Thanks,<br>" + AnkPunchLine + "<BR>​Ankurayan Committee<br>Bandhu, The Friend<br>https://bandhuodisha.in/<br>email: bandhuankurayan@gmail.com<br>Contact:" + BandhuContactNo + "</p>" +
      "</td></tr></table>";
    const options = {
      htmlBody: htmlBody,
      name: "Bandhu Ankurayan"
    };

    try {

      GmailApp.sendEmail(email, subject + AnkVersion, "", options);
      writeLogToSheet([new Date()], rowStart);
    } catch (e) {
      writeLogToSheet(e.message, rowStart);
    }

  }
}
// Write uid to the responses sheet
function writeUidToSheet(uid, rowStart) {
  const ws = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const dataRange = ws.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const indexUid;

  // Locate the index of header UID
  for (const c = 0; c < headers.length; c++) {
    if (headers[c] === uidHeader) {
      indexUid = c;
      break;
    }
  }

  // Insert a new column in C if UID column is not found
  if (indexUid === undefined) {
    ws.insertColumnsAfter(1, 1);
    indexUid = 1;
    ws.getRange("C1").setValue(uidHeader);
    dataRange = ws.getDataRange();
    values = dataRange.getValues();
  }

  // Check UID's in the sheet row by row
  const isNewUid = true;
  for (const r = 1; r < values.length; r++) {
    const currentUid = values[r][indexUid];
    if (currentUid === uid) {
      isNewUid = false;
      break;
    }
  }

  if (isNewUid) {
    ws.getRange(rowStart, indexUid + 1).setValue(uid)
  }
}

// Write Log to the responses sheet
function writeLogToSheet(slog, rowStart) {
  const ws = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const dataRange = ws.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const indexslog;

  // Locate the index of header LOG
  for (const c = 0; c < headers.length; c++) {
    if (headers[c] === sLogHeader) {
      indexslog = c;
      break;
    }
  }

  // Insert a new column in B if LOG column is not found
  if (indexslog === undefined) {
    ws.insertColumnsAfter(1, 1);
    indexslog = 1;
    ws.getRange("B1").setValue(sLogHeader);
    dataRange = ws.getDataRange();
    values = dataRange.getValues();
  }

  // Check LOG in the sheet row by row
  const isNewslog = true;
  for (const r = 1; r < values.length; r++) {
    const currentslog = values[r][indexslog];
    if (currentslog === slog) {
      isNewslog = false;
      break;
    }
  }

  if (isNewslog) {
    ws.getRange(rowStart, indexslog + 1).setValue(slog)
  }
}


// Create a UID in document properties
function createUid(responseId) {
  const docProperties = PropertiesService.getDocumentProperties();
  const key = "uid";
  const id = 1;
  if (docProperties.getProperty(responseId)) {
    const uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      id = parseInt(docProperties.getProperty(key));
    }

    const uid = (Math.pow(10, uidDigitLength) + id).toString().substring(1, uidDigitLength + 1);
    uid = uidPrefix + uid;

    docProperties.setProperty(key, id + 1);
    docProperties.setProperty(responseId, uid);
  }
  return uid;
}

// function a random uid in document properties
function createRandomUid(responseId) {
  const docProperties = PropertiesService.getDocumentProperties();
  const key = "randomUids";
  const ids = [];
  if (docProperties.getProperty(responseId)) {
    const uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      ids = JSON.parse(docProperties.getProperty(key));
    }

    do {
      const randomNumber = Math.pow(10, uidDigitLength) + Math.floor(Math.random() * Math.pow(10, uidDigitLength));
      const uid = randomNumber.toString().substring(1, uidDigitLength + 1);
      uid = uidPrefix + uid;
    } while (ids.indexOf(uid) !== -1)

    ids.push(uid)
    docProperties.setProperty(key, JSON.stringify(ids));
    docProperties.setProperty(responseId, uid);
  }
  return uid;
}
