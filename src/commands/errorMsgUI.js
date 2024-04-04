
function bannerNullHandler(banner){

    if (banner == null) {
        console.log("banner is null");
        const options = {
        height: 30,
        width: 20,
        promptBeforeOpen: false,
    };
        Office.context.ui.displayDialogAsync('https://meg217.github.io/Outlook_Addin_Authorization_Verifier/src/commands/dialog.html', options);
        //difference between errorMessage and informationalMessage?
        mailboxItem.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Please enter a banner marking for this email.",
        });

        //maybe shouln't de-allow event? instead make a dialog box show up? no just makes it stall and say working on request...
        console.log("event should be denied");
        //event.completed({ allowEvent: false });

        event.completed(
        {
            allowEvent: false,
            cancelLabel: "Add a location",
            commandId: "msgComposeOpenPaneButton",
            contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
            errorMessage: "Don't forget to add a meeting location.",
            sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
        }
        );
        
        var errorElement = document.querySelector('div.Dialog1045-title');
        var errorElement2 = document.querySelector('class.ms-Dialog-title title-758');

        console.log(errorElement2);
        console.log(errorElement);


        return;
    }
}