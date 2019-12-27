function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('▼Mail Merge▼')
      .addItem('Open Mail Merge', 'showSidebar')
      .addItem('Clear Sheet', 'clearSheet')
      .addItem('How to Use!', 'showManual')
      .addToUi();
}

function onInstall(e){
  onOpen(e);
};

function showSidebar(){
  var html = HtmlService.createHtmlOutputFromFile('MailMerge').setTitle('Mail Merge by David Sung');
  SpreadsheetApp.getUi().showSidebar(html);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("You can send " + MailApp.getRemainingDailyQuota() + " email today.", "", 3);
}

function showManual() {
  var ui = HtmlService.createHtmlOutputFromFile('UserManual')
      .setWidth(500)
      .setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(ui, 'How to use mail merge?');
}


function clearSheet() { // clear current sheet
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var response = Browser.msgBox('Caution!', 'Delete ' + mySheet.getName() + ' sheet contents?',Browser.Buttons.YES_NO);
  if (response == 'yes') {
    mySheet.getRange(2, 1, mySheet.getMaxRows() - 1, mySheet.getMaxColumns()).clearContent();
  } 
}

function getEmail() {
  return Session.getActiveUser().getEmail();
}

/*********************************************************************************************/
function getDraft() { //get email template from gmail draft
  var draft = [];
  var threads = GmailApp.search('in:draft', 0, 10);  
  if (threads.length === 0) { //if there is no draft.
    var html = HtmlService.createHtmlOutputFromFile('createDraft')
          .setWidth(300)
          .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Email Template');
    return;
  }
  
  for (var i = 0; i < threads.length; i++) {
    draft.push((i+1)+'- '+threads[i].getFirstMessageSubject().substr(0, 40));
  }
  var returnVal = JSON.stringify(draft);
  Logger.log(returnVal);
  return returnVal;
}


/*********************************************************************************************/
function getSenderEmail() { //get sender's email
  var accountEmail = [];
  var UserEmail = Session.getActiveUser().getEmail();
  
  accountEmail.push(UserEmail);
  for (var i = 0; i < GmailApp.getAliases().length; i++) {
    if(GmailApp.getAliases().length > 0){
      accountEmail.push(GmailApp.getAliases()[i]);
    }    
  }
  
  var returnVal = JSON.stringify(accountEmail);
  Logger.log(returnVal);
  return returnVal;
}

/*********************************************************************************************/
function getDataSheet() { //get Mail Merge sheet
  var sheetName = [];  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  for (var i=0;i < sheets.length; i++) {
    sheetName.push(sheets[i].getName());
  }

  var returnVal = JSON.stringify(sheetName);
  Logger.log(returnVal);
  return returnVal;
}


/*********************************************************************************************/
function getResult() { //get Result
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("You can send " + MailApp.getRemainingDailyQuota() + " email today.", "", 1);
  var result = [];
  result.push("<a href='https://www.nshc.net' target='_blank'>NSHC Site</a>" + "<br/>");
  Logger.log(result);
  return result;
}

/*********************************************************************************************/
function getMyName () { //get User Name from Owername of createed Drive file.
  var files = DriveApp.getFilesByName('Mail Merge.txt');
  var fileId;
  Logger.log("1.Mail Merge.txt? : " + files.hasNext());
  
  if (files.hasNext() < 1) {
    // Create a text file with the content "Do not delete this File!"
    files = DriveApp.createFile('Mail Merge.txt', 'Do not delete this File!');
    Logger.log("2.Mail Merge.txt? : " + files);
  }
  
  // get File from ID 
  files = DriveApp.getFilesByName('Mail Merge.txt');
  var fileId = files.next().getId();
  var aFile = DriveApp.getFileById(fileId);
  // Log the names of all users who have edit access to a file.
  var ownerName = aFile.getOwner().getName();
  Logger.log("3.OwnerName : " + ownerName);
  return ownerName;
}


/*********************************************************************************************/
function startMailMerge (draft, sheetName, senderEmail, senderName, ccEmail, bccEmail, isTest) {
  
  try
  { 
  // Check the Draft Start. 
  if (draft.length === 0) { //if there is no draft.
    var html = HtmlService.createHtmlOutputFromFile('createDraft')
          .setWidth(300)
          .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Email Template');
    return;
  }
  // Check the Draft End.
  
 var dateFormat = "MMMM dd, yyyy HH:mm:ss Z zz"; // https://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var dataSheet = ss.getSheetByName(sheetName);
 dataSheet.activate();
 if(dataSheet.getRange(1,dataSheet.getLastColumn()).getValue() != 'Mail Merge Status'){
   dataSheet.getRange(1,dataSheet.getLastColumn()+1).setValue('Mail Merge Status');
 }
 var headers = dataSheet.getRange(1, 1, 1, dataSheet.getLastColumn()).getValues();
 var emailColumnFound = false;
 for(i in headers[0]){
   if(headers[0][i] == "Email Address"){
     emailColumnFound = true;
   }
 }
 if(!emailColumnFound){
   var emailColumn = Browser.inputBox("Which column contains the recipient's email address ? (A, B,...)");
   dataSheet.getRange(emailColumn+''+1).setValue("Email Address");
 }
  
 if (isTest != "test") { // if not test
 ss.toast('Please wait...','Starting Mail Merge!',-1);
 }
  
 var dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn());
 var selectedTemplate = GmailApp.search("in:drafts")[(parseInt(draft.substr(0, 2))-1)].getMessages()[0];
 var emailTemplate = selectedTemplate.getBody();
 var rawContent = selectedTemplate.getRawContent();
 var attachments = selectedTemplate.getAttachments();
 var cc = ccEmail;
 var bcc = bccEmail;
 var selectedAliases = senderName;
 var isTest = isTest;

// Send with inline images by David Sung
 var pattern = /<img.*src="([^"]*)"[^>]*>/; //pattern for img tag
 var matches = pattern.exec(emailTemplate);


 if(emailTemplate.search(/<\img/) != -1){
   var inlineImages = {}; // define the inlineImages
   var imgVars = emailTemplate.match(/<img[^>]+>/g); // find img tags
   
   for(i in imgVars){
     var title = imgVars[i].match(/alt="([^\"]+\")/);
     // Logger.log('title : ' + title);
     var data_surl = imgVars[i].match(/data-surl="([^\"]+\")/);
     if(title != null && data_surl != null){
       data_surl = data_surl[1].substr(4, data_surl[1].length-5); 
       title = title[1].substr(0, title[1].length-1);
       title = title.replace(/(\s*)/g, ""); 
       // Logger.log('title.sub : ' + title);

       for(j in attachments){
         if(attachments[j].getName().replace(/(\s*)/g, "") == title){
           inlineImages[title] = attachments[j].copyBlob().setName(title);
           attachments.splice(j,1);
         }
       }
       
       //var newImg = imgVars[i].replace(/src="[^\"]+\"/,"src=\"cid:" +title+ "\""); //replace the code
       var regEx = new RegExp(data_surl, "gi");       
       var newImg = imgVars[i].replace(regEx, title); //  value.replace(/\-/g,'');
       Logger.log("newImg : " + newImg);
       emailTemplate = emailTemplate.replace(imgVars[i],newImg);
     }
   }
 }
  
//////////////////////////////////////////////////////////////////////////////
  
 objects = getRowsData(dataSheet, dataRange);
 var output = HtmlService.createHtmlOutput().setTitle("Log");
  output.append("<div style='font-size:10pt'><div style='color:#666666;padding:10px;border:1px #efefef solid;background-color:#efefef'>Subject : " + selectedTemplate.getSubject() + "<br/>");
  output.append("Merge Sheet : " + dataSheet.getSheetName() + "<br/>");
  output.append("Sender's Email : " + senderEmail + "<br/>");  
  output.append("Sender's Name : " + senderName + "<br/>");

  if (isTest != "test") {
    if (cc) { output.append("CC : " + cc + "<br/>");}
    if (bcc) {output.append("BCC : " + bcc + "<br/>");}   
  }

  output.append("</div><br/>======= START : " +  Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), dateFormat) + " =======<br/><br/>");

/*********************************************************************************************/
// Function for a Test Email

  if (isTest == "test") {  //if test, sending TEST EMAIL
   ss.toast('Please wait...','Sending a test email!',-1);
    
   var rowData = objects[0];
   var UserEmail = Session.getActiveUser().getEmail(); 
   var emailText = fillInTemplateFromObject(emailTemplate, rowData); 
   var emailSubject = fillInTemplateFromObject(selectedTemplate.getSubject(), rowData);
  var message = {
    htmlBody: emailText,
    subject: "[TEST] " + emailSubject,
    name: senderName,
    from: senderEmail,
    "attachments": attachments,
    "inlineImages": inlineImages
  }
    
  MailApp.sendEmail(UserEmail, "[TEST] " + emailSubject, null, message);
    
    output.append("<span style='color:blue'>Sent test a email to " + UserEmail + "</span><br/>");
    output.append("<br/>======= FINISH : " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), dateFormat) + " =======</div>");
    SpreadsheetApp.getUi().showModalDialog(output, "Test Email Result");    
    ss.toast("Complete! You can send " + MailApp.getRemainingDailyQuota() + " email today.", "", 1);
    return;
 }
 /*********************************************************************************************/ 
  
 for (var i = 0; i < objects.length; ++i) {   
   var rowData = objects[i];
   
   if(rowData.mailMergeStatus != "EMAIL_SENT"){
     
     // Replace markers (for instance {{"First Name"}}) with the 
     // corresponding value in a row object (for instance rowData.firstName).
     
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);     
     var emailSubject = fillInTemplateFromObject(selectedTemplate.getSubject(), rowData);
         
     GmailApp.sendEmail(rowData.emailAddress, emailSubject, null,
                        {name: senderName, attachments: attachments, htmlBody: emailText, cc: cc, bcc: bcc, inlineImages:inlineImages, from : senderEmail});      

     dataSheet.getRange(i+2,dataSheet.getLastColumn()).setValue("EMAIL_SENT");
     output.append ("<span style='color:blue'>Success</span> : " + rowData.emailAddress + " <br/>");
     
     
   } else {
     output.append ("<span style='color:red'>Error</span> : " + rowData.emailAddress + ", <span style='color:gray'>Please empty the 'Mail Merge Status' cell.</span><br/>");
     //ss.toast(rowData.emailAddress + "\'s Mail Merge Status is EMAIL_SENT", "", 3);
   }
   
 }
  output.append("<br/>======= FINISH : " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), dateFormat) + " =======</div>");
  Logger.log(output.getContent());
  SpreadsheetApp.getUi().showModalDialog(output, "Mail Merge Results");
  ss.toast("Complete! You can send " + MailApp.getRemainingDailyQuota() + " email today.", "", 1);
    
    
  }
  
  catch (e)
  {
    Logger.log(e.message);
    SpreadsheetApp.getUi().showModalDialog(e, "Error Message!");
    return e;
  }  
    
    
}


// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance {{"Column name"}}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker {{Column name}}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
 var email = template;
  
 //     https://developers.google.com/apps-script/articles/mail_merge
 //     template.match(/   \$\{\"      [^\"]+  \"  \}/g);   =>  ${"   First Name   "}
  var templateVars = template.match(/\{\{[^\{]+\}\}/g); //(/\$\%[^\%]+\%/g)
 if(templateVars!= null){          
   // Replace variables from the template with the actual values from the data object.
   // If no value is available, replace with the empty string.
   for (var i = 0; i < templateVars.length; ++i) {
     // normalizeHeader ignores ${"} so we can call it directly here.
     var variableData = data[normalizeHeader(templateVars[i])];
     email = email.replace(templateVars[i], variableData || "");
   }
 }
 return email;
}


// This code is reused from the 'Reading Spreadsheet data using JavaScript Objects' tutorial 

function getRowsData(sheet, range, columnHeadersRowIndex) {
 columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
 var numColumns = range.getEndColumn() - range.getColumn() + 1;
 var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
 var headers = headersRange.getValues()[0];
 return getObjects(range.getValues(), normalizeHeaders(headers));
}

function getObjects(data, keys) {
 var objects = [];
 for (var i = 0; i < data.length; ++i) {
   var object = {};
   var hasData = false;
   for (var j = 0; j < data[i].length; ++j) {
     var cellData = data[i][j];
     if (isCellEmpty(cellData)) {
       continue;
     }
     object[keys[j]] = cellData;
     hasData = true;
   }
   if (hasData) {
     objects.push(object);
   }
 }
 return objects;
}

function normalizeHeaders(headers) {
 var keys = [];
 for (var i = 0; i < headers.length; ++i) {
   var key = normalizeHeader(headers[i]);
   if (key.length > 0) {
     keys.push(key);
   }
 }
 return keys;
}

function normalizeHeader(header) {
 var key = "";
 var upperCase = false;
 for (var i = 0; i < header.length; ++i) {
   var letter = header[i];
   if (letter == " " && key.length > 0) {
     upperCase = true;
     continue;
   }
   if (!isAlnum(letter)) {
     continue;
   }
   if (key.length == 0 && isDigit(letter)) {
     continue; // first character must be a letter
   }
   if (upperCase) {
     upperCase = false;
     key += letter.toUpperCase();
   } else {
     key += letter.toLowerCase();
   }
 }
 return key;
}

function isCellEmpty(cellData) {
 return typeof(cellData) == "string" && cellData == "";
}

function isAlnum(char) {
 return char >= 'A' && char <= 'Z' ||
   char >= 'a' && char <= 'z' ||
   isDigit(char);
}

function isDigit(char) {
 return char >= '0' && char <= '9';
}
