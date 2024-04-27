/*
 * OUTLOOK ADDIN TO VERIFY AUTHORIZATION OF USERS AGAINST BANNERS
 */

var mailboxItem;

Office.initialize = function (reason) {
  mailboxItem = Office.context.mailbox.item;
};

/**
 * Handles the OnMessageSend event. Heart of the code.
 * Takes in the onSend event.
 */
function MessageSendVerificationHandler(event) {
  //PROMISE HANDELERS FOR OUTLOOK ITEMS ////////////////////////////////////////
  Promise.all([
    getToRecipientsAsync(),
    getSenderAsync(),
    getBodyAsync(),
    getCCAsync(),
    getBCCAsync(),
  ]).then(([to, sender, body, cc, bcc]) => {
    console.log(
      "\nPROMISE HANDELERS FOR OUTLOOK ITEMS:\nRecipient: " +
        to
          .map(
            (recipient) =>
              recipient.emailAddress + " (" + recipient.displayName + ")"
          )
          .join(", ") +
        "\nCC recipients: " +
        (cc
          ? cc
              .map(
                (recipient) =>
                  recipient.emailAddress + " (" + recipient.displayName + ")"
              )
              .join(", ")
          : "None") +
        "\nBCC recipients: " +
        (bcc
          ? bcc
              .map(
                (recipient) =>
                  recipient.emailAddress + " (" + recipient.displayName + ")"
              )
              .join(", ")
          : "None") +
        "\nSender: " +
        sender.displayName +
        "\nBody: " +
        body
    );

    //BANNER HANDELERS ////////////////////////////////////////////////////////
    console.log("BANNER HANDELERS\n");
    const banner = getBannerFromBody(body);
    bannerNullHandler(banner, event);
    const bannerMarkings = parseBannerMarkings(banner);
    console.log(bannerMarkings.banner);
    if (bannerMarkings.message !== "") {
      errorPopupHandler(bannerMarkings.message, event);
    }

    //CHECK IF AUTHORIZED HANDELERS ////////////////////////////////////////////
    console.log("\nIF AUTHORIZED HANDELERS\n");
    Promise.all([
      //checkRecipientClassification(sender, 'sender', bannerMarkings.banner[0]),
      checkRecipientClassification(to, "to", bannerMarkings.banner[0]),
      checkRecipientClassification(cc, "CC", bannerMarkings.banner[0]),
      checkRecipientClassification(bcc, "BCC", bannerMarkings.banner[0]),
    ]).then(([recipientCheck, ccCheck, bccCheck]) => {
      let authChecksPassed = false;
      let countryChecksPassed = false;
      let hasCheckedCountry = false;

      console.log("LOGGING AUTHORIZATION:");
      console.log("Recipient check: " + recipientCheck);
      console.log("CC check: " + ccCheck);
      console.log("BCC check: " + bccCheck);
   // let message = "";

        let recipient_authorized = recipientCheck[0];
        let recipient_email = recipientCheck[1];
        
        let cc_authorized = ccCheck[0];
        let cc_email = ccCheck[1];
        
        let bcc_authorized = bccCheck[0];
        let bcc_email = bccCheck[1];

        if (!recipient_authorized) {
          errorPopupHandler(`Recipient ${recipient_email} is NOT AUTHORIZED to view this email`, event);
          console.log("Recipient check: not authorized");
        }

        else if (!cc_authorized) {
          errorPopupHandler(`CC'd user ${cc_email} is NOT AUTHORIZED to view this email`, event);
          console.log("CC check: not authorized");
        }
        else if (!bcc_authorized) {
          errorPopupHandler(`BCC'd user ${bcc_email} is NOT AUTHORIZED to view this email`, event);
        } else {
          authChecksPassed = true;
        }

    //  if (!recipientCheck) {
    //    message = "Recipient is NOT AUTHORIZED to view this email";
    //    errorPopupHandler(message, event);
    //  } else if (!ccCheck) {
    //    message = "CC'd user(s) is NOT AUTHORIZED to view this email";
    //    errorPopupHandler(message, event);
    //  } else if (!bccCheck) {
    //    message = "BCC'd user(s) is NOT AUTHORIZED to view this email";
    //    errorPopupHandler(message, event);
    //  } else {
    //    authChecksPassed = true;
    //  }

      //CHECK FOR NOFORN DISSEMINATION HANDELERS ////////////////////////////////////////////
      console.log("\nCHECK FOR NOFORN DISSEMINATION HANDELERS\n");
      dissemination = bannerMarkings.banner[2];

      if (dissemination != null) {
        let dissParts = dissemination.split("/");
        let dissPartsArray = [];
        let hasNOFORN = false;
        for (let i = 0; i < dissParts.length; i++) {
          dissPartsArray.push(dissParts[i]);
          if (dissPartsArray[i] === "NOFORN") {
            hasNOFORN = true;
          }
        }

        if (hasNOFORN === true) {
          //if dissPartsArray does not contain noforn at all then has checked is true
          for (let i = 0; i < dissPartsArray.length; i++) {
            //what other banenrs do we need to check
            //REL TO with the country authorized
            //
            if (dissPartsArray[i] === "NOFORN") {
              //NOFORNEncountered = true;
              Promise.all([
                checkCountryForRecipients("to", to),
                checkCountryForRecipients("CC", cc),
                checkCountryForRecipients("BCC", bcc),
              ]).then(([recipientCheck, ccCheck, bccCheck]) => {
                console.log("LOGGING NOFORN AUTHORIZATION:");
                console.log("To check: " + recipientCheck);
                console.log("CC check: " + ccCheck);
                console.log("BCC check: " + bccCheck);
                let message = "";

                let recipient_authorized = recipientCheck[0];
                let recipient_email = recipientCheck[1];
                
                let cc_authorized = ccCheck[0];
                let cc_email = ccCheck[1];
                
                let bcc_authorized = bccCheck[0];
                let bcc_email = bccCheck[1];
                
            
                    if (!recipient_authorized) {
                      errorPopupHandler(`Recipient ${recipient_email} is a Foreign National and NOT AUTHORIZED to view this email`, event);
                    }
            
                    else if (!cc_authorized) {
                      errorPopupHandler(`CC'd user ${cc_email} is a Foreign National and NOT AUTHORIZED to view this email`, event);
                    }
                    else if (!bcc_authorized) {
                      errorPopupHandler(`BCC'd user ${bcc_email} is a Foreign National and NOT AUTHORIZED to view this email`, event);
                    } else {
                      authChecksPassed = true;
                    }
                
      //          if (!recipientCheck) {
      //            message =
      //              "Recipient is a Foreign National and NOT AUTHORIZED to view this email";
      //            errorPopupHandler(message, event);
      //          } else if (!ccCheck) {
      //            message =
      //              "CC'd user(s) is a Foreign National and NOT AUTHORIZED to view this email";
      //           errorPopupHandler(message, event);
      //          } else if (!bccCheck) {
      //            message =
      //              "BCC'd user(s) is a Foreign National and NOT AUTHORIZED to view this email";
      //            errorPopupHandler(message, event);
      //          } else {
      //            countryChecksPassed = true;
      //          }

                console.log(
                  "Authorization checks passed is: " + authChecksPassed
                );
                console.log("Country checks passed is: " + countryChecksPassed);
                if (countryChecksPassed && authChecksPassed) {
                  event.completed({
                    allowEvent: true,
                  });
                }
              });
            }
          }
        } else {
          //should maybe add something here after gone through loop and didnt find noforn
          console.log(" dissemination is not noforn");
          hasCheckedCountry = true;
        }
      } else {
        console.log(" no disseminations found so dont need to check country");
        hasCheckedCountry = true;
      }

      //AUTH CHECKS PASSED THEN ALLOW EVENT ////////////////////////////////////////////
      console.log("Authorization checks passed is: " + authChecksPassed);
      console.log("hasCheckedCountry is: " + hasCheckedCountry);

      if (authChecksPassed && hasCheckedCountry) {
        event.completed({
          allowEvent: true,
        });
      }
    });
    //./Promise 1
  });
  //./Promise 2
}

/**
 * CHECKS THE CLASSIFICATION LEVEL OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @param {String} documentClassification The classification level of the email
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 * above needs to return not only a boolean but email adress too
 */
function checkRecipientClassification(
  recipients,
  recipientType,
  documentClassification
) {
  // console.log(`Checking ${recipientType} recipients classification`);
  const csvFile =
    "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

  return Promise.all(
    recipients.map((recipient) => {
      const emailAddress = recipient.emailAddress;
      // console.log(`${recipientType} Email Address: ${emailAddress}`);
      if (!emailAddress) {
        // console.log("No recipients for: " + recipientType + " type returned " + recipients.emailAddress);
        return [true, null];
      }
      return userMeetsSecurityClearance(
        csvFile,
        documentClassification,
        emailAddress
      )
        .then((isClearance) => {
          if (!isClearance) {
            console.log(`${emailAddress} is NOT AUTHORIZED to view this email`);
            return [false, emailAddress];
          } else {
            console.log(`${recipientType} is cleared`);
            return [true, emailAddress];
          }
        })
        .catch((error) => {
          console.error(
            `Error while checking ${recipientType} clearance: `,
            error
          );
           return [false, emailAddress];
        });
    })
  ).then((results) => {
    return results; // Return authorization status and email address for each recipient
  });
}

/**
 * CHECKS THE NOFORN STATUS OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 */
function checkCountryForRecipients(recipientType, recipients) {
  // console.log(`Checking ${recipientType} country`);

  const csvFile =
    "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

  return Promise.all(
    recipients.map((recipient) => {
      const emailAddress = recipient.emailAddress;
      // console.log(`${recipientType} Email Address: ${emailAddress}`);
      return check_NOFORN_Access(csvFile, emailAddress)
        .then((isNOFORN) => {
          // console.log(`isNOFORN for ${recipientType} ${emailAddress} returned: ${isNOFORN}`);
          if (!isNOFORN) {
            console.log(
              `${emailAddress} is a Foreign National and NOT AUTHORIZED to view this email`);
            return [false, emailAddress];
          } else {
            console.log(`${recipientType} user is Cleared as USA`);
            return [true, emailAddress];
          }
        })
        .catch((error) => {
          console.error(
            `Error while checking isNOFORN for ${recipientType} ${emailAddress}: `,
            error
          );
          return [false, emailAddress];
        });
    })
  ).then((results) => {
    return results; // Return true if all recipients are cleared
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

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate(
  "MessageSendVerificationHandler",
  MessageSendVerificationHandler
);