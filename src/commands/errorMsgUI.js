function bannerNullHandler(banner, event){

    if (banner == null) {
        console.log("banner is null");
        
        // the below commented out code opens a new window with an html page, maybe need this in the future but not right now
        // const options = {
        // height: 30,
        // width: 20,
        // promptBeforeOpen: false,
        // };
        // Office.context.ui.displayDialogAsync('https://meg217.github.io/Outlook_Addin_Authorization_Verifier/src/commands/dialog.html', options);
        
        
        
        //type can either be errorMessage or informationalMessage
        mailboxItem.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Please enter a banner marking for this email.",
        });

        //event.completed({ allowEvent: false });

        event.completed(
        {
            allowEvent: false,
            cancelLabel: "Ok",
            commandId: "msgComposeOpenPaneButton",
            contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
            errorMessage: "Please enter a banner, banner error detected.",
            //underneath with enable the user to press send anyways, might need later
            //sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
        }
        );


        return;
    }
}