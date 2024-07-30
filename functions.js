function validateBeforeSend(event) {
    const restrictedDomains = ["example.com", "anotherdomain.com"];
    
    Office.context.mailbox.item.to.getAsync({ asyncContext: event }, (result) => {
        const recipients = result.value;
        const asyncContext = result.asyncContext;

        Office.context.mailbox.item.attachments.getAsync((attachmentResult) => {
            const attachments = attachmentResult.value;
            let hasRestrictedDomain = false;

            for (let i = 0; i < recipients.length; i++) {
                const email = recipients[i].emailAddress.toLowerCase();
                for (let j = 0; j < restrictedDomains.length; j++) {
                    if (email.endsWith(restrictedDomains[j])) {
                        hasRestrictedDomain = true;
                        break;
                    }
                }
                if (hasRestrictedDomain) break;
            }

            if (attachments.length > 0 && hasRestrictedDomain) {
                const warningMessage = "Warning: You are sending an email with attachments to restricted domains.";
                Office.context.mailbox.item.notificationMessages.addAsync("attachmentWarning", {
                    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                    message: warningMessage,
                    icon: "icon16",
                    persistent: true
                });

                // Cancel the send event
                asyncContext.completed({ allowEvent: false });
            } else {
                // Allow the send event
                asyncContext.completed({ allowEvent: true });
            }
        });
    });
}