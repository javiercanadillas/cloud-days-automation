function testClean() {
  cleanGSStatus();
}


// TODO Use this for processing exceptions: https://yagisanatode.com/2018/06/10/google-apps-script-getting-input-data-from-a-dialog-box-in-google-sheets/#htmlFile
function addUserToTrack_OneShot() {
  let userid: string;
  let tracklist: string[];
  let userData: any[] = ['','',userid,'','',tracklist];
  addUserToCourses(userData);
}


function testAddUsersToGroup() {
  // Get users data range, and then the values
  let usersDataRange = loadUsersData(firstUserDataRow);
  let usersValues = usersDataRange.getValues();
  addUsersToGroup(usersValues,`cloud-days-${mainConfObj.customerName}@gcprocks.es`);
}

function testCreateGroup() {
  createGroup();
}

function testExportCourse() {
  //let courseID = '42017839495';
  let exportCourseTemplateID = '1b8gWeTpldutXFV_yGGjm_g6zMcz7u1BntPZH16W02Tc';
  listCourses().map(courseID => {
    exportCourseToGDoc(courseID, exportCourseTemplateID);
  });
  //exportCourse(courseID, exportCourseTemplateID);
}

function testListCourses() {
  console.log(listCourses());
}