const DATE_COL = 'C';
const INITIAL_RANGE = "B2:BF";
const MAX_EMAILS = 20;

const DEFAULT_NO = "Unknown";
const DEFAULT_STATUS = "Not Applied";
const DEFAULT_LINK = "Link not Found";
const JOB_ALERT_REGEX = /jobalerts-noreply@linkedin.com|jobs-listings@linkedin.com/;
const JOB_TITLE_REGEX = /Software Engineer(.*)|Software Dev(.*)|Developer(.*)/g;
const JOB_LINK_REGEX = /View Job:\s+https:(.*?)/igs;
const JOB_LINK_END = /-{4,}/;
const JOB_LINK_TFRONT = /View Job:\s+/is;
const HEADER_SLICER = /Your Job Alert for Software Engineer|Top job picks for you/i;
const FOOTER_SLICER = /See all jobs/i;
const TIME_FORMAT = "en-US";
const TIME_ZONE = "America/Los_Angeles";

//Author: Eric Daetz
//Date (Updated): July 2, 2025
//Description: An email parsing script designed to help me populate a Google Sheet with job listings pulled from LinkedIn alert emails.
//             I thought this would be a fun project to do to show that I'm eager to learn while hunting for a job. I wanted to provide as many
//             constants as possible such that this script can be edited easily for future usage or different time zones.
//

//special namespace that executes when document is opened
function onOpen(){

  //Create a sheet object for the spreadsheet this script is linked to
  const sheet = SpreadsheetApp.getActiveSheet();
  
  //get current date
  const currDate = new Date();
  const currDateString = currDate.toLocaleDateString(TIME_FORMAT, {timeZone: TIME_ZONE});

  //Split curr DateTime for comparison
  let currDTArray_strs = currDateString.split('/');
  const currDTArray = currDTArray_strs.map(x => parseInt(x));

  const lastDTArray = getLastDateTime(sheet);

  //before parsing all the emails, make sure it even needs to be updated yet (updates daily)
  if(checkInitialUpdate(sheet, currDTArray,lastDTArray)){
    extractEmails(sheet,lastDTArray);
  }
}

//only run in the initial entry point function
function checkInitialUpdate(sheet, currDTArray,lastDTArray){
  //if there are no entries beyond the header, always update
  const checkedRange = sheet.getRange(INITIAL_RANGE);
  if(checkedRange.isBlank()){
    return true;
  }

  if(checkUpdate(currDTArray,lastDTArray)){
    return true;
  }

  //Logger.log("Initial update not needed");
  return false;
}

//returns TRUE if change needed, FALSE otherwise
function checkUpdate(currDTArray, lastDTArray){
  
  //if lastDTArray is null, always update
  if(lastDTArray == null){
    return true;
  }

  if(currDTArray[2] > lastDTArray[2]){
    return true;
  }
  else if(currDTArray[2] == lastDTArray[2]){
    if(currDTArray[0] == lastDTArray[0] && currDTArray[1] > lastDTArray[1]){
      return true;
    }
    else if(currDTArray[0] > lastDTArray[0]){
      return true;
    }
  }

  Logger.log("Returned false from checkUpdate");
  return false;

}

//Returns TRUE if change needed, FALSE otherwise
function getLastDateTime(sheet){

  const checkedRange = sheet.getRange(INITIAL_RANGE);
  if(checkedRange.isBlank()){
    return null;
  }
  const rowidx = sheet.getLastRow();
  const last_cell_idx = DATE_COL + rowidx;
  const last_cell = sheet.getRange(last_cell_idx);
  const last_cell_value = last_cell.getValue();

  //Google Sheets interprets date format as a Date object and not a string
  let lastDateString = last_cell_value.toLocaleDateString(TIME_FORMAT, {timeZone: TIME_ZONE});

  //Split last DateTime to compare
  let lastDTArray_strs = lastDateString.split('/');
  const lastDTArray = lastDTArray_strs.map(x => parseInt(x));

  return lastDTArray;
}

//Extracts the body of the email
function extractBody(messageBody, emailDateString, sheet){

  //chops off the metadata parts of the message body
  const body_noheader = messageBody.split(HEADER_SLICER)[1];

  //handles cases where format was broken for whatever reason

  if(body_noheader == null){
    Logger.log("Expected syntax not found, must skip this email");
    return;
  }

  const clean_messageBodyArray = body_noheader.split(FOOTER_SLICER);

  if(clean_messageBodyArray[1] == null){
    Logger.log("Expected syntax not found, must skip this email");
    return;
  }

  const clean_messageBody = clean_messageBodyArray[0];
  const extracted_listings = [];
  let valid_job = false;
  let building_str = false;
  let current_idx = 0;

  //extracts job information, excluding links
  const messageBody_lines = clean_messageBody.split("\n");
  for(let i = 0; i < messageBody_lines.length; i++){
    if(JOB_TITLE_REGEX.test(messageBody_lines[i])){
      //push anonymous object to array of listings
      extracted_listings.push({title: messageBody_lines[i], company: messageBody_lines[i+1], location: messageBody_lines[i+2], url: DEFAULT_LINK, no: DEFAULT_NO, date: emailDateString, status: DEFAULT_STATUS});
      current_idx = extracted_listings.length - 1;
      valid_job = true;
    }
    else if(JOB_LINK_REGEX.test(messageBody_lines[i]) && valid_job){
      extracted_listings[current_idx].url = "" + messageBody_lines[i];
      building_str = true;
    }
    else if(building_str){
      //reset flags and stop building the string, append to sheet
      if(JOB_LINK_END.test(messageBody_lines[i])){
        building_str = false;
        valid_job = false;
        extracted_listings[current_idx].url = extracted_listings[current_idx].url.split(JOB_LINK_TFRONT)[1];
        extracted_listings[current_idx].url = extracted_listings[current_idx].url.trimEnd();
        const formatted_url = `=HYPERLINK("${extracted_listings[current_idx].url}", "View Job on LinkedIn")`;
        sheet.appendRow([extracted_listings[current_idx].company, extracted_listings[current_idx].title, extracted_listings[current_idx].date, extracted_listings[current_idx].no, extracted_listings[current_idx].status, formatted_url, extracted_listings[current_idx].location]);
      }
      else{
        extracted_listings[current_idx].url += messageBody_lines[i];
      }
    }

  }
  
}

//Get relevant basic job info from emails
function extractEmails(sheet, lastDTArray){

  Logger.log("Reached extractEmails");

  //Extracting up to 50 Emails; More than that seems unnecessary due to timely job alerts being most relevant
  const threads = GmailApp.getInboxThreads(0,MAX_EMAILS);
  for(let i = MAX_EMAILS - 1; i >= 0; i--){

    const messages = threads[i].getMessages();
    for(let j = messages.length - 1; j >= 0; j--){

      const emailDate = messages[j].getDate();
      const emailDateString = emailDate.toLocaleDateString(TIME_FORMAT, {timeZone: TIME_ZONE});
      const senderAddress = messages[j].getFrom();

      //Split curr DateTime for comparison
      let emailDTArray_strs = emailDateString.split('/');
      const emailDTArray = emailDTArray_strs.map(x => parseInt(x));
      Logger.log(`Sender Address: ${senderAddress}`);
      const verifiedEmail = senderAddress.match(JOB_ALERT_REGEX);

      //If from expected addresses, checks for update necessity
      if(verifiedEmail != null){     
        Logger.log(`Email matched on ${i}, ${j} loop`);
        if(checkUpdate(emailDTArray, lastDTArray)){
          Logger.log("Needs an update");
          extractBody(messages[j].getRawContent(), emailDateString, sheet);
        }
      }
    }
  }

  Logger.log("Email extraction completed.")
  
}
