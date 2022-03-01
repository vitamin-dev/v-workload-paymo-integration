const API_KEY               = 'bb500b90eee8753529740259b0189993';
const API_BASE_URL          = 'https://app.paymoapp.com/api/';
const SHEET_NAME            = 'Paymo API Data';
const PROJECT_ID_COLUMN     = 0;
const PROJECT_CODE_COLUMN   = 1;
const CLIENT_ID_COLUMN      = 2;
const PROJECT_NAME_COLUMN   = 3;
const PROJECT_TIME_COLUMN   = 4;
const PROJECT_BUDGET_COLUMN = 5;
const SHEET_HEADERS         = [
  [
   'Project ID',
   'Project Code',
   'Client ID',
   'Project Name',
   'Project Time',
   'Current Budget'
  ],
];
const DOC_PROP_KEY_NAME     = 'v_updated_meta_key_id';
const DEV_META_KEY_NAME     = 'v_last_updated';
const FIVE_MINUTE_MILLI     = 300000;

7140248 (overarching) + 1431255 (cc social) + 742507 (greater-promise social) = 9314010 mins;
9314010 / 3600 = 2587.225hrs;
3410 (proj budget) - 2587.225 = 822.775 hrs left until budget is met.
284.16666667 (projected budget/month) * 4 (months until renewal) = 1136.66666668 hrs should be left;
1136.66666668 - 822.775 = 822.775 hrs left until budget is met.

/*-------------------------------------*\
  Helpers
\*-------------------------------------*/

/**
 * Returns request parameters for the Paymo API
 *
 * @param { string } type Can be 'GET' or 'POST'. Defaults to 'GET'
 * @return { Object } Object containing HTTP headers
 */
function getPaymoRequestParams_( type = 'GET' ) {
  const password = new Date().toISOString();
  let headers  = {
   'Authorization': 'Basic ' + Utilities.base64Encode(API_KEY + ':' + password),
   'Accept': 'application/json',
  };

  if ( type === 'POST' ) {
   headers = {
   'Authorization': 'Basic ' + Utilities.base64Encode(API_KEY + ':' + password),
   'Accept': 'application/json',
   'Content-type': 'application/json'
   };
  }

  return {
   method: type,
   headers,
   muteHttpExceptions: true,
  };
}

/**
 * Helps search/iterate through existing data range
 *
 * @param { Number }    column      The column # to search in
 * @param { string } searchValue The value to match against the one in the key
 * @return { array } The row of data
 */
function findIn_( column, searchValue ) {
  let results = [];

  const sheet = SpreadsheetApp.getActive().getSheetByName( SHEET_NAME );

  if ( ! sheet ) {
    return;
  }

  const data = sheet.getDataRange().getValues();

  for ( let i = 0; i < data.length; i++ ) {
    if (
    i === 0 ||
    data[i][column] !== searchValue
   ) {
      continue;
    }

    results.push( data[i] );
  }

  return results;
}

/**
 * Rounds time in seconds to rounded hours
 *
 * @param { Number } time Time in seconds.
 * @returns { Number } Time rounded up/down to whole hours
 */
function timeToRoundedHours_( time ) {
  const hours     = time / 3600;
  let roundedHours = Math.floor( hours );

  if ( hours % roundedHours >= .5 ) {
    roundedHours += 1;
  }

  return roundedHours;
}

/**
 * Returns the hours for a given Project ID
 *
 * @param { Number }       projectID The Project ID to find the hours for
 * @return { int|bool } The fround project hours or false on failure
 */
 function findProjectHours_( projectID ) {
  const result = findIn_( PROJECT_ID_COLUMN, projectID );

  if ( ! result.length ) {
    return;
  }

  let projectHours = false;

  if ( result.length === 1 ) {
    projectHours = timeToRoundedHours_( result[0][ PROJECT_TIME_COLUMN ] );
  }

  return projectHours;
}

/**
 * Fetches report object
 *
 * @return { object } The API report object
 */
function fetchReportObject_() {
  const url = API_BASE_URL + 'reports';
  const params   = getPaymoRequestParams_( 'POST' );
  params.payload =  JSON.stringify( {
   'name': `TempReport ${ new Date().toString() }`,
   'type': 'temp',
   'date_interval': 'all_time',
   'clients': 'all_active',
   'projects': 'all',
   'users': 'all',
   'include': {
    'clients': true,
    'projects': true,
   },
   'extra': {
    'display_projects_codes': true,
    'display_projects_budgets': true,
    'display_projects_remaining_budgets': true,
    'order': [
      'clients',
      'projects'
    ],
   },
  } );

  const response = UrlFetchApp.fetch( url, params );

  if ( response.getResponseCode() > 400 ) {
   return;
  }

  return JSON.parse( response.getContentText() );
}

/**
 * Fetch single report object
 * @param { string } projectCode The project code without the #
 * @return { object } The matching report object
 */
function fetchSingleReportObject_( projectCode ) {
  const url = API_BASE_URL + 'projects?where=code=' + projectCode;
  const params   = getPaymoRequestParams_();
  const response = UrlFetchApp.fetch( url, params );

  if ( response.getResponseCode() > 400 ) {
   return;
  }

  const projects = JSON.parse( response.getContentText() ).projects.filter( p => p.code.toUpperCase() === projectCode.toUpperCase() );
  // SpreadsheetApp.getUi().alert( JSON.stringify( JSON.parse( response.getContentText() ).projects ) );
  // let result = projects.length ? projects[0] : false;

  return projects.length ? projects[0] : false;
}

/**
 * Sanitizes spreadsheet input by removing whitespace and
 * possible hash mark if copy-pasted directly from Paymo
 *
 * @param { string } text The text to be sanitized
 * @return { string } The sanitized text
 */
function sanitizeCode_( text ) {
  if ( typeof text !== 'string' ) {
    return null;
  }

  return text.trim().replace( /#/g, '' );
}

function setMetadata_() {
  const document = SpreadsheetApp.getActive();

  // Set metadata w/ last updated date
  let metaKeyId   = PropertiesService
   .getDocumentProperties()
    .getProperty( DOC_PROP_KEY_NAME );
  const updatedOn = new Date().toISOString();

  if ( metaKeyId ) {
   const updatedMeta = document
    .createDeveloperMetadataFinder()
      .withId( Number( metaKeyId ) )
      .find()[ 0 ];
   updatedMeta.setValue( updatedOn );
  } else {
   // Add dev metadata to sheet
   document.addDeveloperMetadata(
    DEV_META_KEY_NAME,
    updatedOn,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT
   );

   // Get metadata that was just added
   const metaKey = document
    .getDeveloperMetadata()
    .filter( m => m.getKey() === DEV_META_KEY_NAME )[0];

   // Get Unique ID from metadata
   metaKeyId = String( metaKey.getId() );

   // Store Unique ID for later retrieval
   PropertiesService
    .getDocumentProperties()
    .setProperty(
      DOC_PROP_KEY_NAME,
      metaKeyId
    );
  }
}

/**
 * Builds the Spreadsheet API sheet and populates data
 * @return { bool } Whether the data was populated or not.
 */
function buildApiSheet_() {
  const report   = fetchReportObject_().reports[0];

  if (
   ! report ||
   report.content.items.length < 1
  ) {
   Browser.msgBox(
    'Paymo Error',
    JSON.stringify( report ),
    Browser.Buttons.OK
   );
   return;
  }

  const document = SpreadsheetApp.getActive();
  let apiSheet = document.getSheetByName( SHEET_NAME );

  if ( apiSheet ) {
   apiSheet.clear();
  } else {
   apiSheet = document.insertSheet( SHEET_NAME, document.getNumSheets()  );
  }

  const reportItems = report.content.items;
  let clientId      = reportItems.filter( item => item.type === 'client' )[0].id;

  let projectRows = [];

  for ( let i = 0; i < reportItems.length; i++ ) {
   if ( reportItems[ i ].type === 'client' ) {
    clientId = reportItems[ i ].level < reportItems[ i + 1 ].level ? reportItems[ i ].id : clientId;
    continue;
   } else if ( reportItems[ i ].type === 'project' ) {
    const project = reportItems[ i ];
    projectRows.push( [
      project.id,
      project.code,
      clientId,
      project.title,
      project.time,
      project.budget_hours
    ] );
   }
  }

  apiSheet.getRange( 1, 1, 1, SHEET_HEADERS[0].length ).setValues( SHEET_HEADERS );
  apiSheet.getRange( 2, 1, projectRows.length, SHEET_HEADERS[0].length )
    .setValues( projectRows );

  // hide sheet
  apiSheet.hideSheet();
}

/**
 * Clears data from sheet and removes metatdata.
 */
function resetApiSheet_() {
  const document = SpreadsheetApp.getActive();
  const apiSheet = document.getSheetByName( SHEET_NAME );

  if ( ! apiSheet ) {
   return;
  }

  // Clear sheet contents
  apiSheet.clear();

  const propertiesStore = PropertiesService.getDocumentProperties();
  const metaKeyId       = propertiesStore.getProperty( DOC_PROP_KEY_NAME );

  if ( ! metaKeyId ) {
   return;
  }

  const updatedMeta = document.createDeveloperMetadataFinder()
    .withId( Number( metaKeyId ) )
    .find()[ 0 ];
  updatedMeta.remove();

  propertiesStore.deleteProperty( DOC_PROP_KEY_NAME);

  SpreadsheetApp.flush();
  Browser.msgBox(
   'Paymo Data Reset',
   'The current hours data has been reset. ' +
   'Please fetch new data by selecting "Update ' +
   'Project Hours Data" from the Paymo menu.',
   Browser.Buttons.OK
  );
}

/**
 * Build process for SS
 */
function apiSheetProcess_() {
  buildApiSheet_();
  setMetadata_();
  updateCustomFunctions_();
  SpreadsheetApp.flush();
}

/**
 * Serves as constructor, initializing sheet and adding data to it.
 */
function initializeSheet_() {
  const apiSheet  = SpreadsheetApp.getActive().getSheetByName( SHEET_NAME );

  if ( ! apiSheet ) {
    apiSheetProcess_();
    return;
  }

  // Check whether sheet needs updating
  const needsUpdate = checkForUpdate_();

  if ( ! needsUpdate ) {
   Browser.msgBox(
    'Paymo Data up to Date',
    'The current hours data was last updated within ' +
    'the previous 5 minutes. Please wait 5 minutes ' +
    'and try again. To force an update right now, ' +
    'first reset the current data from the Paymo menu ' +
    'and try again.',
    Browser.Buttons.OK
   );
  } else {
   apiSheetProcess_();
  }
}

/**
 * Checks wether the current spreadsheet needs an update
 *
 * @returns { bool } True if the spreadsheet data is stale, false otherwise
 */
function checkForUpdate_() {
  const metaKeyId   = PropertiesService
   .getDocumentProperties()
    .getProperty( DOC_PROP_KEY_NAME );

  if ( ! metaKeyId ) {
   return true;
  }

  const updatedMeta = SpreadsheetApp
   .getActive()
   .createDeveloperMetadataFinder()
    .withId( Number( metaKeyId ) )
    .find()[ 0 ];
  const today      = new Date();
  const lastUpdate = new Date( updatedMeta.getValue() );
  lastUpdate.setTime( lastUpdate.getTime() + FIVE_MINUTE_MILLI );

  return (lastUpdate <= today);
}

/**
 * Updates existing cells with custom functions on SS update
 */
function updateCustomFunctions_() {
  const ss = SpreadsheetApp.getActive();
  const customHourCells = ss
   .createTextFinder('=GETRECONCILEDHOURS\\([^)]*\\)')
   .matchFormulaText(true)
   .matchCase(false)
   .useRegularExpression(true)
   .findAll();

  if ( ! customHourCells.length ) {
   return;
  }

  const cellClear = ( c ) => {
   const formula = c.getFormula();
   c.clearContent();
   SpreadsheetApp.flush();
   c.setFormula( formula );
  };

  [ ...customHourCells ].forEach( cellClear );

  const customBudgetHoursCells = ss
   .createTextFinder( '=GETHOURSBUDGET\\([^)]*\\)' )
   .matchFormulaText(true)
   .matchCase(false)
   .useRegularExpression(true)
   .findAll()
  const customTimestampCells = ss
   .createTextFinder('=GETREPORTLASTUPDATED\\([^)]*\\)')
   .matchFormulaText(true)
   .matchCase(false)
   .useRegularExpression(true)
   .findAll();
  const customYearlyCells = ss
   .createTextFinder('=GETHOURSBUDGETFROMYEARLY\\([^)]*\\)')
   .matchFormulaText(true)
   .matchCase(false)
   .useRegularExpression(true)
   .findAll();

  if ( customBudgetHoursCells.length ) {
   [ ...customBudgetHoursCells ].forEach( cellClear );
  }

  if ( customYearlyCells.length ) {
   [ ...customYearlyCells ].forEach( cellClear );
  }

  // Timestamp should be saved last
  if ( customTimestampCells.length ) {
   customTimestampCells.forEach( ( c ) => {
    c.clearContent();
    SpreadsheetApp.flush();
    c.setFormula( '=GETREPORTLASTUPDATED()' );
   } );
  }
}

/*------------ End Helpers---------------*/

/**
 * Connects to Paymo API on open and populates the
 * needed data objects for program execution.
 */
function onOpen() {
  const apiSheet = SpreadsheetApp.getActive().getSheetByName( SHEET_NAME );
  const ssUi     = SpreadsheetApp.getUi();
  const menu     = ssUi.createMenu( 'Paymo' );

  if ( apiSheet && ! apiSheet.isSheetHidden() ) {
   apiSheet.hideSheet();
  }

  menu
    .addItem( 'Update Project Hours Data', 'initializeSheet_' )
    .addToUi();

  menu
    .addItem( 'Reset Paymo Data', 'resetApiSheet_' )
    .addToUi();
}

/**
 * Updates sheet if needed
 */
function updateHelper() {
  const needsUpdate = checkForUpdate_();

  if ( needsUpdate ) {
   apiSheetProcess_();
  }
}

/**
 * Gets timestamp of last report update if available
 *
 * @return { Date|null } Date string if report exists, null otherwise.
 * @customfunction
 */
function getReportLastUpdated() {
  const metaKeyId   = PropertiesService.getDocumentProperties()
    .getProperty( DOC_PROP_KEY_NAME );

  if ( ! metaKeyId ) {
    return null;
  }

  const updatedMeta = SpreadsheetApp
   .getActive()
   .createDeveloperMetadataFinder()
    .withId( Number( metaKeyId ) )
    .find()[ 0 ];

  return Utilities.formatDate(
    new Date( updatedMeta.getValue() ),
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    'MM/dd/YYYY h:mm a',
  );
}

/**
 * Gets reconciled project hours from a project code
 *
 * @param { string } projectCode     A project code to query the projects DB
 * @param { Number }    [untrackedTime] Any additional time not tracked in Paymo to subtract from retainer.
 * @return { int|string } The current retainer hours left in the project or null on failure
 * @customfunction
 */
function getReconciledHours( projectCode, untrackedTime = 0 ) {
  projectCode = sanitizeCode_( projectCode );

  if (
   ! projectCode ||
   typeof projectCode !== 'string'
  ) {
   return null;
  }
  const matchingProjects = findIn_( PROJECT_CODE_COLUMN, projectCode );
  if ( ! matchingProjects.length ) {
   // attempt to fetch directly from API
   const singleReport = fetchSingleReportObject_( projectCode );
    return singleReport ? Number( singleReport.budget_hours ).toPrecision(3) + ' hours remaining' : 'Project not found.';
  }

  const projectTime    = matchingProjects[ 0 ][ PROJECT_TIME_COLUMN ];
  const retainerBudget = Number( matchingProjects[ 0 ][ PROJECT_BUDGET_COLUMN ] );

  return ( Number( retainerBudget ) - timeToRoundedHours_( projectTime ) - untrackedTime ).toPrecision(3) + ' hours remaining';
}

/**
 * Gets current project hours budget from a project code
 *
 * @param { string } projectCode     A project code to query the projects DB
 * @return { int|string } The current retainer hours left in the project or null on failure
 * @customfunction
 */
function getHoursBudget( projectCode ) {
  projectCode = sanitizeCode_( projectCode );

  if (
   ! projectCode ||
   typeof projectCode !== 'string'
  ) {
   return null;
  }

  const matchingProjects = findIn_( PROJECT_CODE_COLUMN, projectCode );
  if ( ! matchingProjects.length ) {
   // attempt to fetch directly from API
   const singleReport = fetchSingleReportObject_( projectCode );

    return singleReport ? singleReport.budget_hours : 'Project not found.';
  }

  const projectBudget = Number( matchingProjects[ 0 ][ PROJECT_BUDGET_COLUMN ] );

  return projectBudget.toPrecision(3);
}

/**
 * Gets current project hours spent from a single project code
 *
 * @param { string } projectCodes A single project code or a csv list of project codes
 * @param { int } yearlyBudget The project's yearly budget
 * @param { date } renewalDate The date the project will renew
 * @returns { int|null } The current retainer hours left in the project for this month or null on failure
 * @customfunction
 */
function getHoursBudgetFromYearly( projectCodes, yearlyBudget, renewalDate ) {
  const sanitizedCodes = projectCodes.split( ',' ).map( sanitizeCode_ );
  const totalTime = sanitizedCodes.reduce( ( acc, code ) => {
    if ( ! code ) {
      return acc;
    }

    const matchingProjects = findIn_( PROJECT_CODE_COLUMN, code );
    if ( ! matchingProjects.length ) {
      // attempt to fetch directly from API
      const singleReport = fetchSingleReportObject_( code );
      return singleReport ? acc + Number( singleReport.budget_hours ).toPrecision(3) : acc;
    }

    return acc + matchingProjects[ 0 ][ PROJECT_TIME_COLUMN ];
  }, 0);

  if (
     totalTime <= 0 ||
     ! yearlyBudget ||
     ! renewalDate
   ) {
   return null;
  }

  const timeLeft        = yearlyBudget - timeToRoundedHours_(totalTime);
  const monthlyBudget   = yearlyBudget / 12;
  const currentMonth    = new Date().getMonth();
  const renewalMonth    = renewalDate.getMonth();
  const remainingMonths = (currentMonth + 1) > renewalMonth ? 12 - currentMonth + renewalMonth : renewalMonth - currentMonth;
  // const timeLeftThisMonth = timeLeft - ( remainingMonths * monthlyBudget );

  return timeLeft - ( remainingMonths * monthlyBudget );
}
