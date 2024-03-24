/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function (reason) {};

/**
 * Handles the OnMessageRecipientsChanged event.
 /**
 * Initializes all of the code.
 * @param {*} event The Office event object
 */
function MessageSendVerificationHandler(event) {
  Promise.all([
    getToRecipientsAsync(),
    getSenderAsync(),
    getBodyAsync()
  ])
  .then(([toRecipients, sender, body]) => {
    console.log("To recipients:");
    toRecipients.forEach(recipient => console.log(recipient.emailAddress));
    console.log("Sender:" + sender.displayName + " " + sender.emailAddress);
    console.log("Body:" + body);
    //const bannerMarkings = parseBannerMarkings(body);
    const messageBodyTest = "TOP SECRET//COMINT-GAMMA/TALENT KEYHOLE//ORIGINATOR CONTROLLED";
    const bannerMarkings = parseBannerMarkings(messageBodyTest);
    console.log(bannerMarkings);

  checkRecipientClassification(toRecipients)
    .then(allowEvent => {
      if (!allowEvent) {
        // Prevent sending the email
        event.completed({ allowEvent: false });
        Office.context.mailbox.item.notificationMessages.addAsync(
          "unauthorizedSending",
          {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: "You are not authorized to send this email to meaganbmueller@gmail.com."
          }
        );
      } else {
        // Allow sending the email
        event.completed({ allowEvent: true });
      }
    })
    .catch(error => {
      console.error("Error occurred while checking recipient classification: " + error);
    });
});
}

/**
 * Gets the 'to' from email.
 */
function getToRecipientsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.to.getAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("unable to get recipients");
        reject("Failed to get To recipients. " + JSON.stringify(result.error));
      } else {
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'sender' from email.
 */
function getSenderAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.from.getAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("unable to get sender");
        reject("Failed to get sender. " + JSON.stringify(result.error));
      } else {
        //const msgFrom = result.value;
        //console.log("Message from: " + msgFrom.displayName + " (" + msgFrom.emailAddress + ")");
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'body' from email.
 */
function getBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("unable to get body");
        reject("Failed to get body. " + JSON.stringify(result.error));
      } else {
        //console.log("this worked");
        resolve(result.value);
      }
    });
  });
}

/**
 * function to parse banner markings
 * parameter is the message body contents
 */
function parseBannerMarkings(body){
  // const cat1_regex = "TOP[\s]*SECRET|TS|(TS)|SECRET|S|(S)|CONFIDENTIAL|C|(C)|UNCLASSIFIED|U|(U)";
  // const cat4_regex = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS";
  // const cat7_regex = "ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  // const cat4_and_cat7 = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  const cat1_regex = /TOP\s*SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U/gi;
  const cat4_regex = /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS/gi;
  const cat7_regex = /ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*USA,\s*AUZ,\s*NZL|REL\s*TO\s*USA,\s*AUS,\s*NZL|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
  const cat4_and_cat7 = /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*USA,\s*AUZ,\s*NZL|REL\s*TO\s*USA,\s*AUS,\s*NZL|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
  
  const Categories = body.split("//");
  const Category_1 = Category1(Categories, cat1_regex);
  const Category_4 = Category4(Categories, cat4_and_cat7, cat7_regex);
  const Category_7 = Category7(Categories, cat7_regex);
  const Together = [Category_1, Category_4, Category_7];
  return Together;
}

function Category1(categories, cat1_regex) {
  if (!categories[0]) {
      console.log("cat one returned null");
      return null;
  } else if (categories[0].toUpperCase().match(cat1_regex)) {
      console.log("returning first category");
      console.log(categories[0].toUpperCase());
      return categories[0].toUpperCase();
  }
  console.log("cat one returned null location 2");
  return null;
}

function Category4(categories, cat4_and_cat7, cat7_regex) {
  if (categories.length < 2) {
      return null;
  } else if (categories[1].toUpperCase().match(cat4_and_cat7)) {
      if (categories[1].toUpperCase().match(cat7_regex)) {
          return categories[1];
      } else {
          return categories[1].toUpperCase();
      }
  }
  return null;
}

function Category7(categories, cat7_regex) {
  if (categories.length < 3) {
      return null;
  } else if (categories[2].toUpperCase().match(cat7_regex)) {
      return categories[2].toUpperCase();
  }
  return null;
}


/**
 * sets session data
 * key and value parameters
 */
function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function(asyncResult) {
      // Handle success or error.
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
      if (value) {
        _tagExternal(value);
      } else {
        _checkForExternal();
      }
    } else {
      console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
      return;
    }
  });
}





//this is the old code from the example
/**
 * Handles the 'to' authentication.
 * @param {*} event The Office event object
 */
function FAKEtoHandler(event) {
  
  Office.context.mailbox.item.to.getAsync(function (asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Failed to get To recipients. " + JSON.stringify(asyncResult.error));
      return;
    }
    
    const toRecipients = asyncResult.value;
    console.log("checking the classification of recipient: "+ toRecipients);
    checkRecipientClassification(toRecipients)
      .then(allowEvent => {
        if (!allowEvent) {
          // Prevent sending the email
          event.completed({ allowEvent: false });
          Office.context.mailbox.item.notificationMessages.addAsync(
            "unauthorizedSending",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
              message: "You are not authorized to send this email to meaganbmueller@gmail.com."
            }
          );
        } else {
          // Allow sending the email
          event.completed({ allowEvent: true });
        }
      })
      .catch(error => {
        console.error("Error occurred while checking recipient classification: " + error);
      });
  });
}

/**
 * Checks the classification level of the recipients.
 * @param {array} recipients The array of recipients
 * @returns {Promise<boolean>} A promise that resolves with a boolean indicating whether the event should proceed
 */
function checkRecipientClassification(recipients) {
  console.log("checkRecipientClassification method"); //debugging

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    
    recipients.forEach(function (recipient) {
      const emailAddress = recipient.emailAddress;
      console.log(emailAddress);

      // Check if recipient is unauthorized
      if (isUnauthorized(emailAddress)) {
        console.log("isUnauthorized returned: " + isUnauthorized(emailAddress));
        allowEvent = false;
      }

    });

    console.log("event should proceed since isUnauthorized returned false");

    // Allow event to proceed if no unauthorized recipient found
    resolve(allowEvent);
  });
}

/**
 * Determines if the recipient is unauthorized.
 * @param {string} emailAddress The recipient's email address
 * @returns {boolean} True if unauthorized, false otherwise
 */
function isUnauthorized(emailAddress) {
  // Check if the recipient's email address matches the unauthorized email address
  return emailAddress === 'meaganbmueller@gmail.com';
}

/**
 * Retrieves the clearance level based on the recipient's email address.
 * @param {string} emailAddress The recipient's email address
 * @returns {string|null} The clearance level required or null if no clearance is needed
 */
function getClearanceLevel(emailAddress) {
  // Perform your logic to determine the clearance level based on the recipient's email address
  // For demonstration, let's assume 'meaganbmueller@gmail.com' requires a 'Classified' clearance
  if (emailAddress === 'meaganbmueller@gmail.com') {
    return 'Classified';
  }
  // If the recipient doesn't require any special clearance, return null
  return null;
}

//  function _setSessionData(key, value) {
//   Office.context.mailbox.item.sessionData.setAsync(
//     key,
//     value.toString(),
//     function(asyncResult) {
//       // Handle success or error.
//       if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//       console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
//       if (value) {
//         _tagExternal(value);
//       } else {
//         _checkForExternal();
//       }
//     } else {
//       console.error(`Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(asyncResult.error)}`);
//       return;
//     }
//   });
// }


// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("MessageSendVerificationHandler", MessageSendVerificationHandler);