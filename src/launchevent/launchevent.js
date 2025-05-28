const customerDomain = "@bizwind.co.jp";

function onMessageSendHandler(event) {
  let externalRecipients = [];

  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;

      recipients.forEach((recipient) => {
        // Extract just the email address, handle potential formatting
        let email = recipient.emailAddress;
        // Remove display name or angle brackets if present (e.g., "Name <email>")
        const match = email.match(/<(.+?)>|[^<>\s]+/);
        email = match ? match[1] || match[0] : email;
        email = email.trim().toLowerCase();
        const domain = customerDomain.toLowerCase();
        
        console.log(`Checking email: ${email}`);
        console.log(`Ends with ${domain}? ${email.endsWith(domain)}`);

        if (!email.endsWith(domain)) {
          externalRecipients.push(email);
        }
      });

      if (externalRecipients.length > 0) {
        console.log(`External recipients found: ${externalRecipients.join(", ")}`);
        event.completed({
          allowEvent: false,
          errorMessage:
            "You are sending this email to external recipients:\n\n" +
            externalRecipients.join("\n") +
            "\n\nAre you sure you want to send it?",
        });
      } else {
        console.log("No external recipients found, allowing send.");
        event.completed({ allowEvent: true });
      }
    } else {
      console.log("Failed to get recipients, allowing send as fallback.");
      event.completed({ allowEvent: true });
    }
  });
}

if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
