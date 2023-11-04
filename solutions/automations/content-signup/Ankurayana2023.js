const sheetName = "FormResponses1"; // replace this sheet name with yours
const uidHeader = "Registration No"; // column name of UID in response sheet
const sLogHeader = "Email Sent Log"; //column name for email log
const uidPrefix = "ANK2023-"; // UID prefeix
const uidDigitLength = 4; // digit length of UID

function _onFormSubmit(e) {
  const { rowStart } = e.range
  var url = SpreadsheetApp.getActive().getFormUrl();
  var form = FormApp.openByUrl(url);
  var response = form.getResponses().pop();
  var url = form.getPublishedUrl();
  const editUrl = response.getEditResponseUrl()
  var responseId = response.getId();

  // for sequential id
  var uid = createUid(responseId);
  // for random id
  //var uid = createRandomUid(responseId);
  writeUidToSheet(uid, rowStart);

  var email = response.getRespondentEmail();
  if (email) {

    var AnkVersion = "Ankurayan 2023"; // Ankurayan Verion
    var AnkContactNo = "9778249509, 9999"; //Contact Numbers
    var AnkPunchLine = "ଅଙ୍କୁରାୟନ ହସର ଫଗୁଣ - ମାଟିରେ ମଗନ"; //Anlrayan Punch Line
    var BandhuContactNo = "9778249509, 9778249509"; //Bandhu office Contact No
    var greet = "Greetings from Bandhu family!!!";
    var title = "Thanks for your registration in ";
    var contmob = "For any enquiry and feedback, you can contact on :  "

    var subject = "Bandhu : Registration Number ";

    var htmlBody = "<p><table style='width: 600px; margin: 10px auto; background: #dfe4ea;'><tr>" +
      "<td style='padding: 0px 20px;'> <font face='verdana'>" + greet + "<br>" + title + AnkVersion +
      "<p>Please refer your Unique Registration No</p>" +
      "<div style='padding: 20px;border-radius: 20px; background: linear-gradient(60deg, #609fd6, #1aafbc);'>" +

      "<div style='font-size: 20px; color: #E06666; text-align: center;'><b>" + uid + "</b></div>" +

      "</div>" + "<br>" + contmob + AnkContactNo +

      "<p>Thanks,<br>" + AnkPunchLine + "<BR>​Ankurayan Committee<br>Bandhu, The Friend<br>https://bandhuodisha.in/<br>email: bandhuankurayan@gmail.com<br>Contact:" + BandhuContactNo + "</p>" +
      "</td></tr></table>";
    var options = {
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
  var ws = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var dataRange = ws.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  var indexUid;

  // Locate the index of header UID
  for (var c = 0; c < headers.length; c++) {
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
  var isNewUid = true;
  for (var r = 1; r < values.length; r++) {
    var currentUid = values[r][indexUid];
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
  var ws = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var dataRange = ws.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];
  var indexslog;

  // Locate the index of header LOG
  for (var c = 0; c < headers.length; c++) {
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
  var isNewslog = true;
  for (var r = 1; r < values.length; r++) {
    var currentslog = values[r][indexslog];
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
  var docProperties = PropertiesService.getDocumentProperties();
  var key = "uid";
  var id = 1;
  if (docProperties.getProperty(responseId)) {
    var uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      id = parseInt(docProperties.getProperty(key));
    }

    var uid = (Math.pow(10, uidDigitLength) + id).toString().substring(1, uidDigitLength + 1);
    uid = uidPrefix + uid;

    docProperties.setProperty(key, id + 1);
    docProperties.setProperty(responseId, uid);
  }
  return uid;
}

// function a random uid in document properties
function createRandomUid(responseId) {
  var docProperties = PropertiesService.getDocumentProperties();
  var key = "randomUids";
  var ids = [];
  if (docProperties.getProperty(responseId)) {
    var uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      ids = JSON.parse(docProperties.getProperty(key));
    }

    do {
      var randomNumber = Math.pow(10, uidDigitLength) + Math.floor(Math.random() * Math.pow(10, uidDigitLength));
      var uid = randomNumber.toString().substring(1, uidDigitLength + 1);
      uid = uidPrefix + uid;
    } while (ids.indexOf(uid) !== -1)

    ids.push(uid)
    docProperties.setProperty(key, JSON.stringify(ids));
    docProperties.setProperty(responseId, uid);
  }
  return uid;
}
