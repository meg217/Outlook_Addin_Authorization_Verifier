// all of the promise functions here such as fetching csv, to, from, and body


/**
 * Gets the csv data file.
 */
function fetchAndParseCSV() {
    const csvUrl =
      "https://meg217.github.io/Outlook_Addin_Authorization_Verifier/assets/accounts.csv";
    console.log("attempting to parse csv");
    return new Promise((resolve, reject) => {
      Papa.parse(csvUrl, {
        download: true,
        header: true,
        complete: (result) => {
          console.log("results are now: " + result.data);
          console.log("or maybe : " + result.data[0]);
          resolve(result.data);
        },
        error: (error) => {
          console.log("error orccured while trying to parse csv");
          reject(error);
        },
      });
    });
  }
  
  /**
   * Gets the 'to' from email.
   */
  function getToRecipientsAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.to.getAsync((result) => {
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
      Office.context.mailbox.item.from.getAsync((result) => {
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
      Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Text,
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log("unable to get body");
            reject("Failed to get body. " + JSON.stringify(result.error));
          } else {
            //console.log("this worked");
            resolve(result.value);
          }
        }
      );
    });
  }

  /**
   * Gets the 'CC' from email.
   */
  function getCCAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.cc.getAsync((result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log("Unable to get CC");
            //reject("Failed to get CC. " + JSON.stringify(result.error));
            cc = null;
            resolve(cc);
          } else {
            resolve(result.value);
          }
        });
    });
  }

  /**
   * Gets the 'BCC' from email.
   */
  function getBCCAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.bcc.getAsync((result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.log("Unable to get BCC");
            //reject("Failed to get BCC. " + JSON.stringify(result.error));
            bcc = null;
            resolve(bcc);
          } else {
            resolve(result.value);
          }
        }
      );
    });
  }