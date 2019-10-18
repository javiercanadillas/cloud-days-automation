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
function contains(a: any[] | string[], obj: any) {
    for (let i = 0; i < a.length; i++) {
        if (a[i] === obj) {
          Logger.log("Found %s", obj);
            return true;
        }
    }
    return false;
}


/**
 *  get HTML Content from Google Doc template
 * @param id Google Doc ID
 */
function _getHTMLContent(id: string) {
  let url = 'https://docs.google.com/feeds/';
  return UrlFetchApp.fetch(url+'download/documents/Export?exportFormat=html&format=html&id='+id).getContentText();
}

/**
 * 
 * @param docID 
 */
function _getGDocAsHTML(docID: string){
  console.log(`Entered _geGDocAsHTML() function. Got docID ${docID}`);
  //needed to get Drive Scope requested and not get a 401 error when running the script
  const forDriveScope = DriveApp.getStorageUsed();
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${docID}&exportFormat=html`;
  // Compose request
  let param = {
        method: "get",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions:true,
  };
  let html = UrlFetchApp.fetch(url,param).getContentText();
  console.log(`_getGDocAsHTML(): Document content in HTML format is ${html}`);
  return html;
}

/**
 * Gets the Google Document ID from the document URL
 * @param url 
 */
function _getIDFromURL(url: string): string{
  let rx = /[-\w]{25,}/;
  console.log(`_getIDFromURL(): got URL ${url} extracted match ${url.match(/[-\w]{25,}/)}`);
  return url.match(rx)[0];
}

/**
 * Loads user information range from main user info sheet
 * @param startRow Starting row, useful for discarding headers
 * @returns Google Sheets range object containing user information
 */
function loadUsersData(startRow: number) {
  // startRow variable if for testing. If not provided, we should really start
  // processing from row 1
  startRow = startRow || 1;
  // Get how far data goes down the sheet
  let dataDepth = SpreadsheetApp.getActive().getSheetByName(provSheetName)
    .getDataRange().getNumRows();
  console.log(`loadUsersData(): loaded data is ${SpreadsheetApp.getActive().getSheetByName(provSheetName).getDataRange().getValues()[dataDepth]}`);
  console.log(`loadUsersData(): data depth is ${dataDepth}`);
    // Get data up to the real email Row
  let dataWidth = idxOf.eMailCheck;
  // Get relevant Data Range as a Range offset of the data range
  return SpreadsheetApp.getActive().getSheetByName(provSheetName)
  .getDataRange().offset(startRow, idxOf.firstName, dataDepth - startRow, dataWidth + 1);
}