/**
 * Create an individual G Suite identity
 * @param userData array containing user data
 * @returns Operation code
 */
function createGSUser(userData: string[]):string {
  // Prepare user creation object
  var userObject = {
    "primaryEmail": userData[idxOf.GSuiteEmail],
    "name": {
      "givenName": userData[idxOf.firstName],
      "familyName": userData[idxOf.lastName]
    },
    "password": userData[idxOf.GSuitePw],
    "changePasswordAtNextLogin": false,
    "orgUnitPath": userData[idxOf.OUPath]
  };

  try {
    AdminDirectory.Users.insert(userObject);
    return 'CREATED';
  } catch (err) {
    console.error('Error for user', userData[idxOf.GSuiteEmail], ': ', err);
    return 'ERR - EXISTS';
  }
  console.log(userObject);
}

/**
 * Removes a given G Suite user from a specific OU
 * @param identity User email address
 * @param oupath Organizational Unit Path as specified in the docs
 */
function deleteGSUser({ identity, oupath }: { identity: string; oupath: string; }) {
  // Substitute error codes accordingly (for example, use booleans)
  let returnCodes = {
    success: 'Removal SUCCESS',
    error: 'Removal ERROR'
  };

  // Using searchQuery to narrow down results
  let options = {
    'primaryEmail': identity,
    'searchQuery': oupath
  };

  try {
    Logger.log("Deleting user %s", identity);
    AdminDirectory.Users.remove(identity);
    Logger.log("User %s removed successfully", identity);
    return returnCodes.success;
  } catch (e) {
    Logger.log("Error deleting user");
    return returnCodes.error;
  }
}

/**
* @description List users matching a specific search query
*
* @param searchQuery String specifying the filter to apply when listing
*        See https://developers.google.com/admin-sdk/directory/v1/guides/search-users
*        for valid queries.
*
* @returns array|null Array containing the users matching
*          the query or null if there aren't any.
**/
function listUsers(searchQuery: string) {
  console.log('listUsers(): Search query is %s', searchQuery);
  let userOptions: object;
  if (searchQuery) {
    userOptions = {
    'maxResults': maxUsersNum,
    'customer': 'my_customer',
    'query': searchQuery
    };
  } else {
    userOptions = {
    'maxResults': maxUsersNum,
    'customer': 'my_customer',
    };
  };
  let page = AdminDirectory.Users.list(userOptions);
  let usersList = page.users;
  if (usersList) {
    return usersList.map(user => user.primaryEmail);
  } else {
    console.error('No matching users found.');
    return null;
  }
}

function createOU(orgUnitPath:string, customerName:string) {
  // REGISTER OU Creation in spreadsheet
  let resource: GoogleAppsScript.AdminDirectory.Schema.OrgUnit = {
    'name': customerName,
    'orgUnitPath': orgUnitPath,
    'description': `Google Cloud Days users for ${customerName}` 
  };
  AdminDirectory.Orgunits.insert(resource, 'my_customer');
  
}

/**
 * Creates a G Suite group to address all users
 */
function createGroup() {
  let customerName = mainConfObj.customerName;
  let resource: GoogleAppsScript.AdminDirectory.Schema.Group = {
    name: `Cloud Days ${customerName}`,
    email: `cloud-days-${customerName}@gcprocks.es`
  };
  AdminDirectory.Groups.insert(resource);
}

/**
 * Add sheet users to group
 * @param usersData 
 * @param groupId 
 */
function addUsersToGroup(usersData: string[][], groupId: string) {
  // iterate users list
  usersData.map(userData => {
    let userEmail = userData[idxOf.GSuiteEmail];
    let userMember = {
      email: userEmail,
      kind: 'admin#directory#member',
      role: 'MEMBER',
      type: 'USER'
    };
    try {
      AdminDirectory.Members.insert(userMember, groupId);
    } catch (e) {
      Logger.log("Error %s trying to add user %s to group %s", e, userEmail, groupId);
    }
  });
}