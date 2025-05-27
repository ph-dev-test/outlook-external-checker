// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const customerDomain = "@bizwind.co.jp";

function onMessageSendHandler(event) {
  let externalRecipients = [];

  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;

      recipients.forEach((recipient) => {
        if (!recipient.emailAddress.includes(customerDomain)) {
          externalRecipients.push(recipient.emailAddress);
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
      // If there's an error retrieving recipients, allow sending to avoid blocking
      event.completed({ allowEvent: true });
    }
  });
}

// Associate the handler with the event
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
