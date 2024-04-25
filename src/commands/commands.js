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
    fetchAndParseCSV(),
    getCCAsync(),
    getBCCAsync(),
  ]).then(([to, sender, body, fetchAndParseCSV, cc, bcc]) => {
    console.log("PROMISE HANDELERS FOR OUTLOOK ITEMS:\nRicipient: " +
      to.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ") +
      "\nCC recipients: " +
      (cc ? cc.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ") : "None") +
      "\nBCC recipients: " +
      (bcc
        ? bcc.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ")
        : "None") +
      "\nSender: " +
      sender.displayName +
      "\nBody: " +
      body);


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
    console.log("IF AUTHORIZED HANDELERS\n");
    Promise.all([
      //checkRecipientClassification(sender, 'sender', bannerMarkings.banner[0]),
      checkRecipientClassification(to, 'to', bannerMarkings.banner[0]),
      checkRecipientClassification(cc, 'CC', bannerMarkings.banner[0]),
      checkRecipientClassification(bcc, 'BCC', bannerMarkings.banner[0])
    ]).then(([recipientCheck, ccCheck, bccCheck]) => {
      console.log("Recipient check: " + recipientCheck);
      console.log("CC check: " + ccCheck);
      console.log("BCC check: " + bccCheck);
      let message = "";
      if (!recipientCheck) {
        console.log("recipient is false so should send message");
        message = "Recipient is NOT AUTHORIZED to view this email";
        errorPopupHandler(message, event);
      } else if (!ccCheck) {
        message = "CC'd user(s) is NOT AUTHORIZED to view this email";
        errorPopupHandler(message, event);
      } else if (!bccCheck) {
        message = "BCC'd user(s) is NOT AUTHORIZED to view this email";
        errorPopupHandler(message, event);
      }
    });
    

    //CHECK FOR NOFORN DISSEMINATION ////////////////////////////////////////////
    console.log("CHECK FOR NOFORN DISSEMINATION\n");
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
          Promise.all([
            checkCountryForRecipients('to', to),
            checkCountryForRecipients('CC', cc),
            checkCountryForRecipients('BCC', bcc)
          ]).then(([recipientCheck, ccCheck, bccCheck]) => {
            console.log("To check: " + recipientCheck);
            console.log("CC check: " + ccCheck);
            console.log("BCC check: " + bccCheck);
          });
        }
      }
    }





  });
}


/**
 * CHECKS THE CLASSIFICATION LEVEL OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @param {String} documentClassification The classification level of the email
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 */
function checkRecipientClassification(recipients, recipientType, documentClassification) {
  console.log(`Checking ${recipientType} recipients classification`);
  const csvFile ="https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";
  
  return Promise.all(recipients.map((recipient) => {
    const emailAddress = recipient.emailAddress;
    console.log(`${recipientType} Email Address: ${emailAddress}`);
    if(!emailAddress){
      console.log("No recipients for: " + recipientType + " type returned " + recipients.emailAddress);
      return true;
    }
    return userMeetsSecurityClearance(csvFile, documentClassification, emailAddress)
      .then((isClearance) => {
        if (!isClearance) {
          console.log(`${emailAddress} is not authorized to view this email`);
          return false;
        } else {
          console.log(`${recipientType} is cleared`);
          return true;
        }
      })
      .catch((error) => {
        console.error(`Error while checking ${recipientType} clearance: `, error);
        return false;
      });
  })).then((results) => {
    return results.every((result) => result); // Return true if all recipients are cleared
  });
}


/**
 * CHECKS THE NOFORN STATUS OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 */
function checkCountryForRecipients(recipientType, recipients) {
  console.log(`Checking ${recipientType} country`);

  const csvFile =
    "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";

  return Promise.all(recipients.map((recipient) => {
    const emailAddress = recipient.emailAddress;
    console.log(`${recipientType} Email Address: ${emailAddress}`);
    return check_NOFORN_Access(csvFile, emailAddress)
      .then((isNOFORN) => {
        console.log(`isNOFORN for ${recipientType} ${emailAddress} returned: ${isNOFORN}`);
        if (!isNOFORN) {
          console.log(`${emailAddress} is a Foreign National and not authorized to view this email`);
          return false;
        } else {
          console.log(`${recipientType} user is Cleared as USA`);
          return true;
        }
      })
      .catch((error) => {
        console.error(`Error while checking isNOFORN for ${recipientType} ${emailAddress}: `, error);
        return false;
      });
  })).then((results) => {
    return results.every((result) => result); // Return true if all recipients are cleared
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
