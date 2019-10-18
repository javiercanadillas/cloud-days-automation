/**
 * Cleans array from empty elements
 * @param array to be cleaned
 */
function cleanArray(array: any[]) {
  return array.filter(function (el) {
    if (el[0] != '') {
      return el
    }
  })
}

/**
 * Converts a 2D array into an object literal
 *   Example: [ [1,'a'], [2,'b'], [3,'c'] ] to { 1: "a", 2: "b", 3: "c" } 
 * @param array 2d array. Columns are key properties, corresponding
 *  rows are values.
 * @returns object
 */
function objectify(array: any[][]) : object {
    return array.reduce(function(result, currentArray) {
        result[currentArray[0]] = currentArray[1];
        return result;
    }, {})
}

/**
 * Checks if a string is empty using RegEx
 * @param str 
 */
function isBlank(str: string) {
    return (!str || /^\s*$/.test(str));
}

/**
 * Checks for value
 * @param a object
 * @param obj 
 */
function contains(a, obj) {
    for (let i = 0; i < a.length; i++) {
        if (a[i] === obj) {
          Logger.log("Found %s", obj);
            return true;
        }
    }
    return false;
}

/**
 * Cleans the status information from G Suite or Classroom columns
 * @param index The column index to be cleaned
 */
function cleanStatusColumn(index: number) {
  let usersDataRange = loadUsersData(firstUserDataRow);
  let limit = usersDataRange.getNumRows();
  usersDataRange.offset(0,index,limit,1).clearContent();
}
/**
 * Sets/Unsets toggle checks for G Suite or Classroom actions
 * @param toggle Decides between G Suite or Classroom
 */
function toggleChecks({ toggle, index }: { toggle: boolean; index: number; }) {
  let usersDataRange = loadUsersData(firstUserDataRow);
  let limit = usersDataRange.getNumRows();
  usersDataRange.offset(0,index,limit,1).setValue(toggle);  
}

/**
 * Loads user information range from main user info sheet
 * @param {number} startRow Starting row, useful for discarding headers
 * @returns {Range} Google Sheets range object containing user information
 */
function loadUsersData(startRow: number) {
  // startRow variable if for testing. If not provided, we should really start
  // processing from row 1
  startRow = startRow || 1;
  // Get how far data goes down the sheet
  let dataDepth = SpreadsheetApp.getActive().getSheetByName(provSheetName)
    .getDataRange().getNumRows();
  // Get data up to the real email Row
  let dataWidth = idxOf.realEmail;
  // Get relevant Data Range as a Range offset of the data range
  return SpreadsheetApp.getActive().getSheetByName(provSheetName)
  .getDataRange().offset(startRow, idxOf.firstName, dataDepth - startRow, dataWidth);
}

/**
 * Creates a Drive folder from a given path
 * @param fullPathToFolder string containing the folder path
 */
function createFolderFromPathName(fullPathToFolder:string): GoogleAppsScript.Drive.Folder {
  // Split the different elements of the path
  let paths = fullPathToFolder.split("/");
  console.log(paths);
  let curFolder = DriveApp.getRootFolder();
  console.log(`Root folder is ${curFolder}`);
  cleanArray(paths).map(element => {
    let folder = curFolder.getFoldersByName(element);
    console.log(`Folder is now ${element}`);
    curFolder = folder.hasNext() ? folder.next() : curFolder.createFolder(element);
  });
  return curFolder;
}