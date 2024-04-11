/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var mailboxItem;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
};

/**
 * Handles the OnMessageSend event.
 /**
 * Initializes all of the code.\
 * @param {*} event The Office event object
 */
function MessageSendVerificationHandler(event) {
  //promise is to encapsulate all the async functions
  Promise.all([
    getToRecipientsAsync(),
    getSenderAsync(),
    getBodyAsync(),
    fetchAndParseCSV(),
    getCCAsync(),
    getBCCAsync(),
  ]).then(([toRecipients, sender, body, fetchAndParseCSV, cc, bcc]) => {
    console.log("To recipients:");
    toRecipients.forEach((recipient) => console.log(recipient.emailAddress));
    console.log("Sender:" + sender.displayName + " " + sender.emailAddress);
    console.log("CC: " + cc.emailAddress);
    console.log("BCC: " + bcc.emailAddress);
    console.log("Body:" + body);
    const banner = getBannerFromBody(body);

    // Check if the banner is null error
    bannerNullHandler(banner, event);

    //const messageBodyTest = "TOP SECRET//COMINT-GAMMA/TALENT KEYHOLE//ORIGINATOR CONTROLLED";
    const bannerMarkings = parseBannerMarkings(banner);
    console.log(bannerMarkings);

    checkRecipientClassification(toRecipients,bannerMarkings[0])
      .then((allowEvent) => {
        if (!allowEvent) {
          // Prevent sending the email
          console.log("Prevent sending email");
          event.completed({ allowEvent: false });
          Office.context.mailbox.item.notificationMessages.addAsync(
            "unauthorizedSending",
            {
              type: Office.MailboxEnums.ItemNotificationMessageType
                .ErrorMessage,
              message: "You are not authorized to send this email",
            },
            (result) => {
              console.log(result);
            }
          );
        } else {
          // Allow sending the email
          event.completed({ allowEvent: true });
        }
      })
      .catch((error) => {
        console.error(
          "Error occurred while checking recipient classification: " + error
        );
      });
  });
}

function fetchCSVData(url) {
  return fetch(url).then((csvData) => parseCSV(csvData));
}

/**
 * sets session data
 * key and value parameters
 */
function _setSessionData(key, value) {
  Office.context.mailbox.item.sessionData.setAsync(
    key,
    value.toString(),
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`sessionData.setAsync(${key}) to ${value} succeeded`);
        if (value) {
          _tagExternal(value);
        } else {
          _checkForExternal();
        }
      } else {
        console.error(
          `Failed to set ${key} sessionData to ${value}. Error: ${JSON.stringify(
            asyncResult.error
          )}`
        );
        return;
      }
    }
  );
}

/**
 * Checks the classification level of the recipients.
 * @param {array} recipients An array of recipients
 * @param {String} documentClassication The classication level of the email dictated by category 1 of banner
 * @returns {Promise<boolean>} Returns true if all recipients are permitted to view the contents of the email
 */
function checkRecipientClassification(recipients,documentClassification) {
  console.log("checkRecipientClassification method"); //debugging
  //userMeetsSecurityClearance(filePath, documentClassification, email) {

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    const csvFile = "./assets.users.csv";

    // If a single recipient is not permitted, the entire send fails
    recipients.forEach(function (recipient) {
      const emailAddress = recipient.emailAddress;
      if (userMeetsSecurityClearance(csvFile,documentClassification,emailAddress)) {
        console.log(emailAddress + " is not authorized to view this email");
        allowEvent = false;
      }
    });

    resolve(allowEvent);
  });

  // Old Method
  /**
   * return new Promise((resolve, reject) => {
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
  });*/
}

/**
 * Determines if the recipient is unauthorized.
 * @param {string} emailAddress The recipient's email address
 * @returns {boolean} True if unauthorized, false otherwise
 */
function isUnauthorized(emailAddress) {
  // Check if the recipient's email address matches the unauthorized email address
  return emailAddress === "meaganbmueller@gmail.com";
}

/**
 * Retrieves the clearance level based on the recipient's email address.
 * @param {string} emailAddress The recipient's email address
 * @returns {string|null} The clearance level required or null if no clearance is needed
 */
function getClearanceLevel(emailAddress) {
  // Perform your logic to determine the clearance level based on the recipient's email address
  // For demonstration, let's assume 'meaganbmueller@gmail.com' requires a 'Classified' clearance
  if (emailAddress === "meaganbmueller@gmail.com") {
    return "Classified";
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
Office.actions.associate(
  "MessageSendVerificationHandler",
  MessageSendVerificationHandler
);
