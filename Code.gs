function main() {

  var spreadsheetID = "1hUkrpG5CktIfQY8nuzyGVAE44Pzm0QSwv8G6frk1V1c";
  var emailTemplateID = "15cExqOqofM-WO-g7gMToLt-WtmDXTuy0xo8S4ALwppM";
  var birthdayImage = 'https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQ4Oww2LuWUNidsUpcoRkOo3Vtu4c8q33mIPQ&s';
  var ccEmails = '';
  
  birthdayImageBlob = UrlFetchApp.fetch(birthdayImage).getBlob();


  var spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  var sheet = spreadsheet.getActiveSheet();
  var document = DocumentApp.openById(emailTemplateID);

  // Check if the sheet exists
  if (!sheet) {
    Logger.log("Sheet 'Birthdays' not found.");
    return;
  }

  if (!document) {
    Logger.log("document 'email template' not found.");
    return;
  }

  const emailBody = document.getBody();

  // Get the last row in the sheet that has data.
  var numRows = sheet.getLastRow();

  // Load data in the first two columns from the second row till the last row. 
  // Remember: The first row has column headers so we don’t want to load it.
  var range = sheet.getRange(2, 1, numRows - 1, 5).getValues();

  // Use a for loop to process each row of data
  for (var index in range) {
    // For each row, get the person’s name and their birthday
    var row = range[index];
    var name = row[0];
    var birthday = row[1];
    var email = row[2];
    var project = row[3];
    var photo = row[4];

    // Check if the person’s birthday is today
    if (isBirthdayToday(birthday)) {
      var emailContent = emailBody.getText();
      var personPicture = false;
      if (photo) {
        personPicture = getImageByDriveUrl(photo);
      }
      emailContent = emailContent.replaceAll('<NAME>',name);
      emailContent = emailContent.replaceAll('<PROJECT>',project);
      emailContent = emailContent + '<p><img src="cid:birthday_image"/></p>';
      if (personPicture) {
        emailContent = emailContent + '<p><img src="cid:'+ personPicture.getId() +'"/></p>';
      }
      emailReminder(name, email, emailContent, ccEmails, personPicture);
    }
  }
}

// Check if a person’s birthday is today
function isBirthdayToday(birthday) {
  // If birthday is a string, convert it to date
  if (typeof birthday === "string")
    birthday = new Date(birthday);
  var today = new Date();
  if ((today.getDate() === birthday.getDate()) &&
      (today.getMonth() === birthday.getMonth())) {
    return true;
  } else {
    return false;
  }
}

// Function to send the email reminder
function emailReminder(name, bdayemail, emailTemplate, ccEmails, personPicture) {
  var subject = "Happy Birthday " + name;
  var recipient = Session.getActiveUser().getEmail() + ',' + bdayemail;
  var body = "" + emailTemplate + "";
  var inlineImages = {};
  inlineImages['birthday_image'] = birthdayImageBlob;
  if (personPicture) {
    inlineImages[personPicture.getId()] = personPicture.getBlob();
  }
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: body,
    cc: ccEmails,
    inlineImages: personPicture ?  inlineImages : {}
  });
}

function getImageByDriveUrl(driveUrl) {
  try {
    const fileId = extractFileIdFromUrl(driveUrl);
    const file = DriveApp.getFileById(fileId);
    return file;
  } catch (error) {
    Logger.log(`Error fetching image: ${error.message}`);
    return null;
  }
}

function extractFileIdFromUrl(url) {
  // Regular expression to extract the file ID from the URL
  var regex = /[-\w]{25,}/;
  var matches = url.match(regex);
  if (matches && matches.length > 0) {
    return matches[0];
  } else {
    throw new Error('Invalid Google Drive URL');
  }
}
