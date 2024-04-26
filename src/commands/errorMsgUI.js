//all of the error handlers with popup messages will be handled here

function bannerNullHandler(banner, event) {
  if (banner == null) {
    console.log("banner is null");

    //type can either be errorMessage or informationalMessage
    // mailboxItem.notificationMessages.addAsync("NoSend", {
    // type: "errorMessage",
    // message: "Please enter a banner marking for this email.",
    // });

    //event.completed({ allowEvent: false });

    event.completed({
      allowEvent: false,
      cancelLabel: "Ok",
      commandId: "msgComposeOpenPaneButton",
      contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
      errorMessage: "Please enter a banner, banner error detected.",
      //     //underneath with enable the user to press send anyways, might need later
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });

    return;
  }
}

//CHANGE:
function errorPopupHandler(errorMsg, event) {
  console.log(
    "error detected, handeling popup and setting allow event to false"
  );
  event.completed({
    allowEvent: false,
    cancelLabel: "Ok",
    commandId: "msgComposeOpenPaneButton",
    contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
    errorMessage: errorMsg,
    //sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
  });
  return;
}
