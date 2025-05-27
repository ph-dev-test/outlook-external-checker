/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const customerDomain = "@bizwind.co.jp";

function onMessageSendHandler(event) {
  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      // Log all recipient details for debugging
      console.log("Recipients found:", recipients);
      console.log("Raw email addresses:", recipients.map(r => r.emailAddress));

      const externalRecipients = recipients
        .filter(recipient => {
          let email = recipient.emailAddress || "";
          console.log("Raw value for recipient:", email);
          
          // Clean the email: trim, convert to lowercase, and extract from potential display name format
          email = email.trim().toLowerCase();
          // Handle formats like "Display Name <email@domain.com>"
          const match = email.match(/<([^>]+)>/);
          if (match) {
            email = match[1];
          }
          // Fallback: if no angle brackets, try to extract the email part
          else if (email.includes("@")) {
            email = email.split(" ").pop(); // Take the last part, assuming email is after display name
          }
          
          const isExternal = !email.endsWith(customerDomain.toLowerCase());
          console.log(`Processed email: ${email}, isExternal: ${isExternal}`);
          return isExternal;
        })
        .map(recipient => recipient.emailAddress);

      if (externalRecipients.length > 0) {
        console.log("External recipients detected:", externalRecipients);
        event.completed({
          allowEvent: false,
          errorMessage:
            "The mail includes some external recipients, are you sure you want to send it?\n\n" +
            externalRecipients.join("\n") +
            "\n\nClick Send to send the mail anyway.",
        });
      } else {
        console.log("No external recipients, allowing send");
        event.completed({ allowEvent: true });
      }
    } else {
      // Failed to get recipients, cancel send for safety and log error
      console.error("Failed to get recipients: " + asyncResult.error.message);
      event.completed({ allowEvent: false });
    }
  });
}

// Map the event handler for Outlook on Windows, Mac, and other platforms
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
