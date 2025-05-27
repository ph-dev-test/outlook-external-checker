/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const customerDomain = "@bizwind.co.jp";

function onMessageSendHandler(event) {
  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      // Log full recipient objects and raw email addresses
      console.log("Full recipient objects:", JSON.stringify(recipients, null, 2));
      console.log("Raw email addresses:", recipients.map(r => r.emailAddress));

      const externalRecipients = recipients
        .filter(recipient => {
          // Default to empty string if emailAddress is undefined or null
          let rawEmail = recipient.emailAddress || "";
          console.log("Raw email value:", rawEmail);

          // Clean the email: trim and convert to lowercase
          let email = rawEmail.trim().toLowerCase();

          // Step 1: Handle "Display Name <email@domain.com>" format
          let match = email.match(/<([^>]+)>/);
          if (match) {
            email = match[1];
            console.log("Extracted from angle brackets:", email);
          }

          // Step 2: Fallback - split by spaces and find part with @
          if (!match && email.includes("@")) {
            const parts = email.split(" ");
            email = parts.find(part => part.includes("@")) || email;
            console.log("Extracted from split:", email);
          }

          // Step 3: Final check - ensure we have a valid email-like string
          if (!email.includes("@")) {
            console.log("Warning: No valid email format detected for:", rawEmail);
            return false; // Treat invalid format as internal to avoid false popup
          }

          // Check if email ends with customer domain
          const isExternal = !email.endsWith(customerDomain.toLowerCase());
          console.log(`Final processed email: ${email}, isExternal: ${isExternal}`);
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
