/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const customerDomain = "@your-domain.com";

function onMessageSendHandler(event) {
  let externalRecipients = [];
  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      var recipients = asyncResult.value;
      recipients.forEach((recipient) => {
        if (!recipient.emailAddress.includes(customerDomain)) {
          externalRecipients.push(recipient.emailAddress);
        }

        if (externalRecipients.length > 0) {
          event.completed({
            allowEvent: false,
            errorMessage:
              "The mail includes some external recipients, are you sure you want to send it?\n\n" +
              externalRecipients.join("\n") +
              "\n\nClick Send to send the mail anyway.",
          });
        } else {
          event.completed({ allowEvent: true });
        }
      });
    } else {
      event.completed({ allowEvent: true });
    }
  });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
