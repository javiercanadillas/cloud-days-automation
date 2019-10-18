// BEGIN GLOBALS

// Main user info sheet name
const provSheetName = 'Provisioning';
// Named range containing track to ID configuration object
const classesByID = 'classesByID';
// Named range containing main configuration object
const mainConfig = 'mainConfig';

// Store indexes of Provisioning sheet header for convenience
const idxOf = {
  "firstName": 0,
  "lastName": 1,
  "GSuiteEmail": 2,
  "GSuitePw": 3,
  "OUPath": 4,
  "track": 5,
  "GSuiteCheck": 6,
  "CRCheck": 7,
  "GSStatus": 8,
  "CRStatus": 9,
  "realEmail": 10,
}

// Initialize main configuration object
const mainConfObj: {
  customerName?: string
  domain?: string
  password?: string
  orgUnitPath?: string
  gOrgUnitPath?: string
  mainExportFolder?: string
  courseTemplateID?: number
} = objectify(SpreadsheetApp.getActive().getRangeByName(mainConfig)
                            .getValues());

// Set First Data row for main user info sheet
const firstUserDataRow = 1;
// Cap number of processed users, when searching for users in G Suite
const maxUsersNum = 150;

// Get [trackName,courseID} array
const trackToIDArray = SpreadsheetApp.getActiveSpreadsheet()
  .getRangeByName(classesByID).getValues();
// Get track: CourseID object
const trackToID = objectify(cleanArray(trackToIDArray));

// END GLOBALS

/**
 * Renders Spreadsheet menu with options
 */
function onOpen() {
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addSubMenu(ui.createMenu('Config Actions')
        .addItem('Create OU', 'createGSuiteOU')
        .addItem('Create Group', 'createGroup'))
      .addSeparator()
      .addSubMenu(ui.createMenu('G Suite Actions')
        .addItem('Create', 'addGSUsers')
        .addItem('Check', 'checkGSUsers')
        .addItem('Remove', 'deleteGSUsers')
        .addItem('Clean Status', 'cleanGSStatus')
        .addItem('Set All Checks', 'setAllGSChecks')
        .addItem('Unset All Checks', 'unsetAllGSChecks'))
      .addSeparator()
      .addSubMenu(ui.createMenu('Classroom Actions')
        .addItem('Add to classes', 'addCRUsers')
        .addItem('Check membership', 'checkCRUsers')
        .addItem('Remove membership', 'removeCRUsers')
        .addItem('Clean Status', 'cleanCRStatus')
        .addItem('Set All Checks', 'setAllCRChecks')
        .addItem('Unset All Checks', 'unsetAllCRChecks'))
      .addToUi();
}

// BEGIN SHAMEFUL WRAPPERS FOR MENUS

/**
 * Creates a OU for the customer identities
 */
function createGSuiteOU() {
  createOU(mainConfObj.orgUnitPath,'my_customer');
}

/**
 * Adds selected users to G Suite domain
 */
function addGSUsers() {
  processUsers({ gsAction: 'ADD', crAction: '' });
}

/**
 * Checks G Suite domain status for selectec users
 */
function checkGSUsers() {
  processUsers({ gsAction: 'CHECK', crAction: '' });
}

/**
 * Removes selected users from G Suite domain
 */
function deleteGSUsers() {
  processUsers({ gsAction: 'REMOVE', crAction: '' });
}

/**
 * Subscribes selected users to Classroom courses
 */
function addCRUsers() {
  processUsers({ gsAction: '', crAction: 'ADD' });
}

/**
 * Checks classroom courses subscription status for selected users
 */
function checkCRUsers() {
  processUsers({ gsAction: '', crAction: 'CHECK' });
}

/**
 * Removes classroom couirses subscription for selected users
 */
function removeCRUsers() { 
  processUsers({ gsAction: '', crAction: 'REMOVE' });
}

/**
 * Cleans G Suite status column in Main User Info sheet
 */
function cleanGSStatus() { 
  cleanStatusColumn(idxOf.GSStatus);
}

/**
 * Cleans Classroom status column in Main User Info sheet
 */
function cleanCRStatus() {
  cleanStatusColumn(idxOf.CRStatus);
}

/**
 * Sets all checks in G Suite Provisioning column
 */
function setAllGSChecks() { 
  toggleChecks({ toggle: true, index: idxOf.GSuiteCheck });
}

/**
 * Unsets all checks in G Suite Provisioning column
 */
function unsetAllGSChecks() {
  toggleChecks({ toggle: false, index: idxOf.GSuiteCheck });
}

/**
 * Sets all checks in Classroom Provisioning column
 */
function setAllCRChecks() {
  toggleChecks({ toggle: true, index: idxOf.CRCheck });
}

/**
 * Unsets all checks in Classroom Provisioning column
 */
function unsetAllCRChecks() {
  toggleChecks({ toggle: false, index: idxOf.CRCheck });
} 
// END SHAMEFUL WRAPPERS FOR MENUS

/**
 * Processes G Suite Users from the Provision sheet that are checked
 *    and meet the criteria and/or register the in courses if needed.
 * @param gsAction  action to be performed on G Suite. Valid
 *    values are 'ADD', 'REMOVE' and 'CHECK'
 * @param crAction action to be performed on Classroom. Valid
 *    values are 'ADD', 'REMOVE' and 'CHECK'
 */
function processUsers({ gsAction, crAction }: { gsAction: string; crAction: string; }) {

  console.info('processUsers(): function entered');
  // Prepare array to store G Suite action result
  let gsStatus : string[][];
  // Prepare array to store Classroom action result
  let crStatus : string[][];

  // Get users data range, and then the values
  let usersDataRange = loadUsersData(firstUserDataRow);
  let usersValues = usersDataRange.getValues();
  // Get status columns ranges so we can update with operation results
  let GSStatusRange = usersDataRange.offset(0,idxOf.GSStatus,usersValues.length,1);
  let CRStatusRange = usersDataRange.offset(0,idxOf.CRStatus,usersValues.length,1);
  
  // Detect type of action
  switch (gsAction) {
    case 'ADD':
      gsStatus = usersValues.map(user => {
        // Check all required attributes are there and if user is
        // marked for processing
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && !(isBlank(user[idxOf.firstName]))
          && !(isBlank(user[idxOf.lastName]))
          && user[idxOf.GSuiteCheck]) { //user marked for provisioning action
          // All is good, go create the user
          console.log('Trying to create identity for user ', user[idxOf.GSuiteEmail]);
          return [createGSUser(user)];
        } else {
          return ['SKIPPED'];
        }
      });
      // Update all results at once, to reduce access to Spreadsheet
      GSStatusRange.setValues(gsStatus);
      // We're done, exit the switch statement
      break;

    case 'CHECK':
      console.info('processUsers(): Checking exiting G Suite identities');
      // First, get list of existing identities in the OU (including Googlers)
      let searchQuery:string = `orgUnitPath=${mainConfObj.orgUnitPath}`;
      let gSearchQuery:string = `orgUnitPath=${mainConfObj.gOrgUnitPath}`;
      // Get array containing existing users
      let existingUsers = listUsers(searchQuery).concat(listUsers(gSearchQuery));
      gsStatus = usersValues.map(user => {
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && user[idxOf.GSuiteCheck]) { //user marked for provisioning action
          console.log('processUsers(): Checking identity ', user[idxOf.GSuiteEmail]);
          if (contains(existingUsers, user[idxOf.GSuiteEmail].toLowerCase())) {
            return ['EXISTS'];
          } else {
            return ['MISSING'];
          }
        } else {
          return ['SKIPPED'];
        }
      });
      GSStatusRange.setValues(gsStatus);
      break;

    case 'REMOVE':
      gsStatus = usersValues.map(user => {
        // Check all required attributes are there and if user is
        // marked for processing
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && user[idxOf.GSuiteCheck]) { //user marked for provisioning action
          console.log('Trying to delete identity ', user[idxOf.GSuiteEmail]);
          return [deleteGSUser({ identity: user[idxOf.GSuiteEmail], oupath: searchQuery })];
        } else {
          return ['SKIPPED'];
        }
      });
      GSStatusRange.setValues(gsStatus);
      break;

    default:
      break;
  }

  switch (crAction) {  
    case 'ADD': {
      crStatus = usersValues.map((user) => {
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && user[idxOf.CRCheck]) { // used marked form classroom processing
          // All is good, go add the user to the courses
          console.log('Trying to perform classroom registration for user'
            , user[idxOf.GSuiteEmail]);
          return [addUserToCourses(user)];
        } else {
          return ['SKIPPED'];
        }
      });
      CRStatusRange.setValues(crStatus)
      break;
    }
    case 'CHECK': {
      let courseStudents = listStudents();  
      crStatus = usersValues.map(user => {
        // Check all required attributes are there and if user
        // belongs to courses
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && user[idxOf.CRCheck]) { // used marked form classroom processing
          // All is good, go check courses for the user
          console.log('processUsers(): Checking courses registration for user %s'
            , user[idxOf.GSuiteEmail]);
          return [checkCourses(user[idxOf.GSuiteEmail], courseStudents)];
        } else {
          return ['SKIPPED'];
        }
      });
    }
      CRStatusRange.setValues(crStatus);
      break;
    case 'REMOVE':
      let courseStudents = listStudents();  
      crStatus = usersValues.map(user => {
        if (!(isBlank(user[idxOf.GSuiteEmail]))
          && user[idxOf.CRCheck]) { // used marked form classroom processing
          // All is good, go remove the user to the courses
          console.log('Trying to perform classroom registration for user'
            , user[idxOf.GSuiteEmail]);
          return [removeUserFromCourses(user, courseStudents)];
        } else {
          return ['SKIPPED'];
        }
      });
      CRStatusRange.setValues(crStatus);
      break;
    default:
      break;
  }
}