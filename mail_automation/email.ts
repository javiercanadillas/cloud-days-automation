//@TODO Improve Docs to HTML conversion -- https://gist.github.com/simonw/0acc8b879787ee30ddfdc5c4d9998e5d

// GLOBALS
// Sheet containing main user info
const provSheetName = 'Mailing';
// Range containing mail configuration object
const eMailConfig = 'eMailConfig';
// Range containing track to ID configuration object
const classesByID = 'classesByID';
const mainConfig = 'mainConfig';

// Store indexes of Provisioning sheet header for convenience
const idxOf = {
  "firstName": 0,
  "lastName": 1,
  "GSuiteEmail": 2,
  "GSuitePw": 3,
  "track": 4,
  "realEmail": 5,
  "eMailCheck": 6,
  "eMailStatus": 7
}

// List properties we're expecting from the email configuration object
interface emailConfObj {
  sendMode?: string
  eMailContentDoc?: string
  subject?: string
  from?: string
  replyTo?: string
  cc?: string
  attachment?: Blob
  nameDetector?: string
  trackDetector?: string
  GSuiteUserDetector?: string
  GSuitePwDetector?: string
}


/**
 * Render actions menu 
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
  .addItem('Process e-mails', 'sendEmail')
      .addToUi();
}


/**
 * Main function to send e-mails to customers
 * @param startRow Row where we start processing users for e-mail sending
 */
function sendEmail(startRow: number) {
  
  // startRow variable if for testing. If not provided, we should really start
  // processing from row 1
  startRow = startRow || 1;
  console.log(`sendEmail(): starting processing e-mails from row ${startRow}`);
  // Load users data
  let usersDataRange = loadUsersData(startRow);
  console.log(`sendEmail(): Loaded users data is ${usersDataRange.getValues()}`);
  let usersData = cleanArray(usersDataRange.getValues());
  // Load eMail configuration data
  let eMailConfObj: emailConfObj = objectify(SpreadsheetApp.getActive().getRangeByName(eMailConfig)
                             .getValues());
  // Prepare a data range to store e-mail sending processing results
  let eMailStatusRange = usersDataRange.offset(0,idxOf.eMailStatus,usersData.length,1);
  
  // Load email content as HTML
  let content = _getGDocAsHTML(_getIDFromURL(eMailConfObj.eMailContentDoc));
  
  let sendStatus: string[] = [];

  eMailStatusRange.setValues(
    usersData.map(userData => {
      let emailAddress: string = userData[idxOf.realEmail];
      let sendDate = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy');
      console.log(`sendEmail(): Processing for user ${emailAddress}`);
      if(!isBlank(emailAddress) && userData[idxOf.eMailCheck]) {
        let name = userData[idxOf.firstName];
        let emailBody = content.replace(eMailConfObj.nameDetector,name)
        .replace(eMailConfObj.trackDetector,userData[idxOf.track])
              .replace(eMailConfObj.GSuiteUserDetector,userData[idxOf.GSuiteEmail])
              .replace(eMailConfObj.GSuitePwDetector,userData[idxOf.GSuitePw]);
        try {
          switch (eMailConfObj.sendMode) {
            case 'DRAFT':
              GmailApp.createDraft(
                emailAddress,
                eMailConfObj.subject,
                '',
                {
                  'cc': eMailConfObj.cc,
                  'htmlBody': emailBody
                });
              console.log(`sendEmail(): DRAFTED e-mail for user ${emailAddress}`);
              return [`DRAFTED at ${sendDate}`];
              break;
            case 'FINAL':
              GmailApp.sendEmail(
                emailAddress,
                eMailConfObj.subject,
                '',
                {
                  'cc': eMailConfObj.cc,
                  'htmlBody': emailBody
                });
              console.log(`sendEmail(): SENT e-mail for user ${emailAddress}`);
              return [`SENT at ${sendDate}`];
              break;
            default:
              break;
          }
        } catch(e) {
          return [`ERROR at ${sendDate}`];
          console.log(`sendEmail(): Error processing e-mail for user ${emailAddress}: ${e}`);
        }
      } else {
          console.log(`sendEmail(): SKIPPED e-mail for user ${emailAddress}`);
          return [`SKIPPED at ${sendDate}`];
      }
    })
  );
}