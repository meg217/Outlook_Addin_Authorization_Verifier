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
    console.log(
      "To recipients: " +
        toRecipients.forEach((recipient) => console.log(recipient.emailAddress))
    );

    console.log("Sender:" + sender.emailAddress);
    console.log("CC: " + cc.emailAddress);
    console.log("BCC: " + bcc.emailAddress);
    console.log("Body:" + body);
    const banner = getBannerFromBody(body);

    // Check if the banner is null error
    bannerNullHandler(banner, event);

    //const messageBodyTest = "TOP SECRET//COMINT-GAMMA/TALENT KEYHOLE//ORIGINATOR CONTROLLED";
    const bannerMarkings = parseBannerMarkings(banner);
    //CHANGE
    console.log(bannerMarkings.banner);

    //CHANGE
    if (bannerMarkings.message !== "") {
      errorPopupHandler(bannerMarkings.message, event);
    }

    //CHANGE
    //fix this first!!!
    const recipientCheck = checkRecipientClassification(toRecipients, bannerMarkings.banner[0], event);
    const senderCheck = checkSenderClassification(sender, bannerMarkings.banner[0], event);
    const ccCheck = check_CC_Classification(cc, bannerMarkings.banner[0], event);
    const bccCheck = check_BCC_Classification(bcc, bannerMarkings.banner[0], event);

    //need to find a way for if cc and bcc are not null then check them
    if (recipientCheck && senderCheck){
      console.log("recipient and sender cleared. recipient check returned " + recipientCheck.resolve + " and senderCheck returned "+ senderCheck.resolve);
      event.completed({
        allowEvent: true,
      });
    }


    dissemination = bannerMarkings.banner[2];

    if (dissemination != null) {
      let dissParts = dissemination.split("/");
      let dissPartsArray = [];

      for (let i = 0; i < dissParts.length; i++) {
        dissPartsArray.push(dissParts[i]);
      }
      for (let i = 0; i < dissPartsArray.length; i++) {
        if (dissPartsArray[i] === "NOFORN") {
          //NOFORNEncountered = true;
          const RecipientMsgreturn = checkRecipientCountry(toRecipients, event);
          console.log("Function checkRecipientCountry returned: " + RecipientMsgreturn);
          const CCMsgreturn = check_CC_Country(cc, event);
          console.log("Function check_CC_Country returned: " + CCMsgreturn);
          const BCCMsgreturn = check_BCC_Country(bcc, event);
          console.log("Function check_BCC_Country returned: " + BCCMsgreturn);
        }
      }
    }
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
 * Checks the classification level of the sender
 * @param {array} sender The sender
 * @param {String} documentClassication The classication level of the email dictated by category 1 of banner
 * @returns {Promise<boolean>} Returns true the sender is permitted to view the information they are sending
 */
function checkSenderClassification(
  sender,
  documentClassification,
  event
) {
  console.log("checkSenderClassification method"); //debugging
  console.log("checkSenderClass - Sender: " + sender);
  console.log(
    "checkSenderClass - Sender: " + documentClassification
  );

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a sender is not permitted, the send fails
      const emailAddress = sender.emailAddress;
      console.log("Sender Email Address: " + emailAddress);
      userMeetsSecurityClearance(csvFile, documentClassification, emailAddress)
        .then((isClearance) => {
          console.log("is clearence returned: " + isClearance);
          if (!isClearance) {
            console.log(emailAddress + " to send information contained in this email.");
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage: "Sender is NOT AUTHORIZED to send information contained in this email.",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("Sender is Cleared");
            return true;
            // event.completed({
            //   allowEvent: true,
            // });
          }
        })
        .catch((error) => {
          console.error("Error while checking isClearance: ", error);
        });
    resolve(allowEvent);
  });
}

/**
 * Checks the classification level of the recipients.
 * @param {array} recipients An array of recipients
 * @param {String} documentClassication The classication level of the email dictated by category 1 of banner
 * @returns {Promise<boolean>} Returns true if all recipients are permitted to view the contents of the email
 */
function checkRecipientClassification(
  recipients,
  documentClassification,
  event
) {
  console.log("checkRecipientClassification method"); //debugging
  //userMeetsSecurityClearance(filePath, documentClassification, email) {
  console.log("checkRecipientClass - Recipient: " + recipients);
  console.log(
    "checkRecipientClass - Classification: " + documentClassification
  );

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    //KEVIN - Changed "./assets.users.csv" to "./assets.accounts.csv"
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a single recipient is not permitted, the entire send fails
    for (const recipient of recipients) {
      const emailAddress = recipient.emailAddress;
      console.log("Recipient Email Address: " + emailAddress);
      userMeetsSecurityClearance(csvFile, documentClassification, emailAddress)
        .then((isClearance) => {
          console.log("is clearence returned: " + isClearance);
          if (!isClearance) {
            console.log(emailAddress + " is not authorized to view this email");
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage: "Recipient is NOT AUTHORIZED to see this email.",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("Recipient is Cleared");
            return true;
            // event.completed({
            //   allowEvent: true,
            // });
          }
        })
        .catch((error) => {
          console.error("Error while checking isClearance: ", error);
        });
    }
    resolve(allowEvent);
  });
}

function checkRecipientCountry(recipients, event) {
  console.log("checkRecipientCountry Function");

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    //KEVIN - Changed "./assets.users.csv" to "./assets.accounts.csv"
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a single recipient is not permitted, the entire send fails
    for (const recipient of recipients) {
      const emailAddress = recipient.emailAddress;
      console.log("Recipient Email Address: " + emailAddress);
      check_NOFORN_Access(csvFile, emailAddress)
        .then((isNOFORN) => {
          console.log("isNOFORN returned: " + isNOFORN);
          if (!isNOFORN) {
            console.log(
              emailAddress +
                " is a Foreign National and not authorized to view this email"
            );
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage:
                "Recipient is NOT AUTHORIZED to see this email: NOT RELEASABLE TO FOREIGN NATIONALS",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("Recipient is Cleared as USA");
            event.completed({
              allowEvent: true,
            });
          }
        })
        .catch((error) => {
          console.error("Error while checking isNOFORN: ", error);
        });
    }
    resolve(allowEvent);
  });
}

/**
 * Checks the classification level of the users CCed.
 * @param {array} CCs An array of people who were CC
 * @param {String} documentClassication The classication level of the email dictated by category 1 of banner
 * @returns {Promise<boolean>} Returns true if all users who are CCed are permitted to view the contents of the email
 */
function check_CC_Classification(
  CCs,
  documentClassification,
  event
) {
  console.log("check_CC_Classification method"); //debugging
  console.log("checkCCClass - CC: " + CCs);
  console.log(
    "checkCCClass - Classification: " + documentClassification
  );

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a cced user is not permitted, the send fails
    for (const cc of CCs) {
      const emailAddress = cc.emailAddress;
      console.log("CC Email Address: " + emailAddress);
      userMeetsSecurityClearance(csvFile, documentClassification, emailAddress)
        .then((isClearance) => {
          console.log("is clearence returned: " + isClearance);
          if (!isClearance) {
            console.log(emailAddress + " is not authorized to view this email");
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage: "CCed user is NOT AUTHORIZED to see this email.",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("CCed user is Cleared");
            return true;
            // event.completed({
            //   allowEvent: true,
            // });
          }
        })
        .catch((error) => {
          console.error("Error while checking isClearance: ", error);
        });
    }
    resolve(allowEvent);
  });
}

function check_CC_Country(CCs, event) {
  console.log("checkRecipientCountry Function");

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    //KEVIN - Changed "./assets.users.csv" to "./assets.accounts.csv"
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a cced user is not permitted, the entire send fails
    for (const cc of CCs) {
      const emailAddress = cc.emailAddress;
      console.log("CC Email Address: " + emailAddress);
      check_NOFORN_Access(csvFile, emailAddress)
        .then((isNOFORN) => {
          console.log("isNOFORN returned: " + isNOFORN);
          if (!isNOFORN) {
            console.log(
              emailAddress +
                " is a Foreign National and not authorized to view this email"
            );
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage:
                "CCed user is NOT AUTHORIZED to see this email: NOT RELEASABLE TO FOREIGN NATIONALS",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("CCed user is Cleared as USA");
            event.completed({
              allowEvent: true,
            });
          }
        })
        .catch((error) => {
          console.error("Error while checking isNOFORN: ", error);
        });
    }
    resolve(allowEvent);
  });
}

/**
 * Checks the classification level of the users CCed.
 * @param {array} BCCs An array of people who were CC
 * @param {String} documentClassication The classication level of the email dictated by category 1 of banner
 * @returns {Promise<boolean>} Returns true if all users who are CCed are permitted to view the contents of the email
 */
function check_BCC_Classification(
  BCCs,
  documentClassification,
  event
) {
  console.log("check_BCC_Classification method"); //debugging
  console.log("checkBCCClass - BCC: " + BCCs);
  console.log(
    "checkBCCClass - Classification: " + documentClassification
  );

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a cced user is not permitted, the send fails
    for (const bcc of BCCs) {
      const emailAddress = bcc.emailAddress;
      console.log("BCC Email Address: " + emailAddress);
      userMeetsSecurityClearance(csvFile, documentClassification, emailAddress)
        .then((isClearance) => {
          console.log("is clearence returned: " + isClearance);
          if (!isClearance) {
            console.log(emailAddress + " is not authorized to view this email");
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage: "BCCed user is NOT AUTHORIZED to see this email.",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("BCCed user is Cleared");
            return true;
            // event.completed({
            //   allowEvent: true,
            // });
          }
        })
        .catch((error) => {
          console.error("Error while checking isClearance: ", error);
        });
    }
    resolve(allowEvent);
  });
}

function check_BCC_Country(BCCs, event) {
  console.log("checkRecipientCountry Function");

  return new Promise((resolve, reject) => {
    let allowEvent = true;
    //KEVIN - Changed "./assets.users.csv" to "./assets.accounts.csv"
    const csvFile =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

    // If a cced user is not permitted, the entire send fails
    for (const bcc of BCCs) {
      const emailAddress = bcc.emailAddress;
      console.log("BCC Email Address: " + emailAddress);
      check_NOFORN_Access(csvFile, emailAddress)
        .then((isNOFORN) => {
          console.log("isNOFORN returned: " + isNOFORN);
          if (!isNOFORN) {
            console.log(
              emailAddress +
                " is a Foreign National and not authorized to view this email"
            );
            event.completed({
              allowEvent: false,
              cancelLabel: "Ok",
              commandId: "msgComposeOpenPaneButton",
              contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
              errorMessage:
                "BCCed user is NOT AUTHORIZED to see this email: NOT RELEASABLE TO FOREIGN NATIONALS",
              sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
            });
          } else {
            console.log("BCCed user is Cleared as USA");
            event.completed({
              allowEvent: true,
            });
          }
        })
        .catch((error) => {
          console.error("Error while checking isNOFORN: ", error);
        });
    }
    resolve(allowEvent);
  });
}

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
