// @ts-check

const sheetName = "FormResponses1"; // replace this sheet name with yours
const uidHeader = "Registration No"; // column name of UID in response sheet
const sLogHeader = "Email Sent Log"; //column name for email log
const uidPrefix = "ANK2023-"; // UID prefeix
const uidDigitLength = 4; // digit length of UID
const EMAIL_TEMPLATE_DOC_URL = 'https://docs.google.com/document/d/1m86GT2-4M1o8Nn-32Hg-RFoSmrjE0zyXVou6ds_cNzs/edit?usp=sharing';
const EMAIL_SUBJECT = 'Ankurayan 2023';

// Ankurayan details
const AnkVersion = "Ankurayan 2023"; // Ankurayan Verion
const AnkContactNo = "9778249509, 9999"; //Contact Numbers
const AnkPunchLine = "ଅଙ୍କୁରାୟନ ହସର ଫଗୁଣ - ମାଟିରେ ମଗନ"; //Anlrayan Punch Line
const BandhuContactNo = "9778249509, 9778249509"; //Bandhu office Contact No
const greet = "Greetings from Bandhu family!!!";
const title = "Thanks for your registration in ";
const contmob = "For any enquiry and feedback, you can contact on :  "
const subject = "Bandhu : Registration Number ";


// Keys:
const classStd = "Class"
const emailAddress = "Email Address"
const fatherSNameପିତାଙ୍କନାମ = "Father's Name (ପିତାଙ୍କ ନାମ)"
const homeAddressଘରଠିକଣାAt = "Home Address, (ଘର ଠିକଣା) At:"
const homeAddressଘରଠିକଣାDistrict = "Home Address, (ଘର ଠିକଣା) District:"
const homeAddressଘରଠିକଣାPin = "Home Address, (ଘର ଠିକଣା) PIN:"
const homeAddressଘରଠିକଣାPost = "Home Address, (ଘର ଠିକଣା) Post:"
const mobileWhatsApp = "Mobile (WhatsApp)"
const offlineLocation = "Offline Location"
const participantNameପ୍ରତିଯୋଗୀଙ୍କନାମ = "Participant Name (ପ୍ରତିଯୋଗୀ ଙ୍କ ନାମ)"
const schoolAddressବିଦ୍ୟାଳୟଠିକଣାLankapada = "School Address (ବିଦ୍ୟାଳୟ ଠିକଣା) Lankapada"
const schoolAddressବିଦ୍ୟାଳୟଠିକଣାSarankul = "School Address (ବିଦ୍ୟାଳୟ ଠିକଣା) Sarankul"
const schoolNameInCaseSchoolNameNotFoundInAboveListବିଦ୍ୟାଳୟରନାମLankapada = "School Name (in case school name not found in above list) (ବିଦ୍ୟାଳୟ ର ନାମ  ) Lankapada"
const schoolNameInCaseSchoolNameNotFoundInAboveListବିଦ୍ୟାଳୟରନାମSarankul = "School Name (in case school name not found in above list) (ବିଦ୍ୟାଳୟ ର ନାମ  ) Sarankul"
const schoolNameବିଦ୍ୟାଳୟରନାମLankapada = "School Name (ବିଦ୍ୟାଳୟର ନାମ) - Lankapada"
const schoolNameବିଦ୍ୟାଳୟରନାମSarankul = "School Name (ବିଦ୍ୟାଳୟର ନାମ) - Sarankul"
const selectOneOrMoreActivities = "Select one or more activities"
const timestamp = "Timestamp"
let valueOf; // function which returns the value of given key. undefined, if the key is not found
let uid;

function _onFormSubmit(e) {
  console.log("Got Arguments: ", JSON.stringify(e));
  const responses = e.namedValues;
  valueOf = key => responses[key].join().trim();

  const { rowStart } = e.range
  const url = SpreadsheetApp.getActive().getFormUrl();
  console.log({url});
  // exception from below hence using triggerUid - which is the same id from the given user
  // const form = FormApp.openByUrl(url);
  // const response = form.getResponses().pop();
  const responseId = e.triggerUid ;

  // for sequential id
  uid = createUid(responseId);
  // for random id
  //const uid = createRandomUid(responseId);
  writeUidToSheet(uid, rowStart);

  const email = valueOf(emailAddress);
  if (email) {
    let status = 'not sent';
    try {

      MailApp.sendEmail({
        to: email,
        subject: EMAIL_SUBJECT,
        htmlBody: createEmailBody(),
      });

      status = 'Sent';
    } catch (e) {
      writeLogToSheet(e.message, rowStart);
    }

    // Append the status on the spreadsheet to the responses' row.
    let sheet = SpreadsheetApp.getActiveSheet();
    let row = sheet.getActiveRange().getRow();
    let column = e.values.length + 1;
    sheet.getRange(row, column).setValue(status);

    console.log("status=" + status + "; responses=" + JSON.stringify(responses));

  }
}

/**
 * Creates email body and includes the links based on topic.
 *
 * @param {string} recipient - The recipient's email address.
 * @param {string[]} topics - List of topics to include in the email body.
 * @return {string} - The email body as an HTML string.
 */
function createEmailBody() {
  const name = valueOf(participantNameପ୍ରତିଯୋଗୀଙ୍କନାମ);

  let topicsHtml = topics.map(function(topic) {
  let url = topicUrls[topic];
    return '<li><a href="' + url + '">' + topic + '</a></li>';
  }).join('');
  topicsHtml = '<ul>' + topicsHtml + '</ul>';

  // Make sure to update the emailTemplateDocId at the top.
  let docId = DocumentApp.openByUrl(EMAIL_TEMPLATE_DOC_URL).getId();
  let emailBody = docToHtml(docId);
  emailBody = emailBody.replace(/{{NAME}}/g, name);
  emailBody = emailBody.replace(/{{TOPICS}}/g, topicsSelected());
  emailBody = emailBody.replace(/{{REGISTRATION_MESSAGE_HTML}}/g, registrationMessageHtml);

  return emailBody;
}

const topicsSelected = () => {
    const keys = [
      mobileWhatsApp,
      classStd,
      fatherSNameପିତାଙ୍କନାମ,
      homeAddressଘରଠିକଣାAt,
      homeAddressଘରଠିକଣାDistrict,
      homeAddressଘରଠିକଣାPin,
      homeAddressଘରଠିକଣାPost,
      offlineLocation,
      schoolAddressବିଦ୍ୟାଳୟଠିକଣାLankapada,
      schoolAddressବିଦ୍ୟାଳୟଠିକଣାSarankul,
      schoolNameInCaseSchoolNameNotFoundInAboveListବିଦ୍ୟାଳୟରନାମLankapada,
      schoolNameInCaseSchoolNameNotFoundInAboveListବିଦ୍ୟାଳୟରନାମSarankul,
      schoolNameବିଦ୍ୟାଳୟରନାମLankapada,
      schoolNameବିଦ୍ୟାଳୟରନାମSarankul,
    ];

    let selections = "<h3>You have selected:</h3>";

    keys.forEach( key => {
      const value = valueOf(key);
      if(value) selections += `<h4>${key}</h4> <p>${value}</p><br/>`;
    });
    selections += `<h4>Activities</h4> <p>${valueOf(selectOneOrMoreActivities)}</p><br/>`;

    return selections;
}

const registrationMessageHtml =  "<p><table style='width: 600px; margin: 10px auto; background: #dfe4ea;'><tr>" +
"<td style='padding: 0px 20px;'> <font face='verdana'>" + greet + "<br>" + title + AnkVersion +
"<p>Please refer your Unique Registration No</p>" +
"<div style='padding: 20px;border-radius: 20px; background: linear-gradient(60deg, #609fd6, #1aafbc);'>" +

"<div style='font-size: 20px; color: #E06666; text-align: center;'><b>" + uid + "</b></div>" +

"</div>" + "<br>" + contmob + AnkContactNo +

"<p>Thanks,<br>" + AnkPunchLine + "<BR>​Ankurayan Committee<br>Bandhu, The Friend<br>https://bandhuodisha.in/<br>email: bandhuankurayan@gmail.com<br>Contact:" + BandhuContactNo + "</p>" +
"</td></tr></table>";

// Write uid to the responses sheet
function writeUidToSheet(uid, rowStart) {
  const ws = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const dataRange = ws.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  let indexUid;

  // Locate the index of header UID
  for (let c = 0; c < headers.length; c++) {
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
  let isNewUid = true;
  for (let r = 1; r < values.length; r++) {
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
  let   indexslog;

  // Locate the index of header LOG
  for (let c = 0; c < headers.length; c++) {
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
  let isNewslog = true;
  for (let r = 1; r < values.length; r++) {
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
  let uid;
  let id = 1;
  if (docProperties.getProperty(responseId)) {
    uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      id = parseInt(docProperties.getProperty(key));
    }

    let uid = (Math.pow(10, uidDigitLength) + id).toString().substring(1, uidDigitLength + 1);
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
  let uid;
  if (docProperties.getProperty(responseId)) {
    uid = docProperties.getProperty(responseId);
  } else {
    if (docProperties.getProperty(key)) {
      ids = JSON.parse(docProperties.getProperty(key));
    }

    do {
      const randomNumber = Math.pow(10, uidDigitLength) + Math.floor(Math.random() * Math.pow(10, uidDigitLength));
      uid = randomNumber.toString().substring(1, uidDigitLength + 1);
      uid = uidPrefix + uid;
    } while (ids.indexOf(uid) !== -1)

    ids.push(uid)
    docProperties.setProperty(key, JSON.stringify(ids));
    docProperties.setProperty(responseId, uid);
  }
  return uid;
}
