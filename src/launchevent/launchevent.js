const customerDomain = "@bizwind.co.jp";

function onMessageSendHandler(event) {
  let externalRecipients = [];

  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;

      recipients.forEach((recipient) => {
        const email = recipient.emailAddress.trim().toLowerCase();
        if (!email.endsWith(customerDomain)) {
          externalRecipients.push(email);
        }
      });

      if (externalRecipients.length > 0) {
        event.completed({
          allowEvent: false,
          errorMessage:
            "You are sending this email to external recipients:\n\n" +
            externalRecipients.join("\n") +
            "\n\nAre you sure you want to send it?",
        });
      } else {
        event.completed({ allowEvent: true });
      }
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
