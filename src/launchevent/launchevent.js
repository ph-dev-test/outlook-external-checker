// Define the customer domain
const customerDomain = "@bizwind.co.jp";

// Handler for the OnMessageSend event
function onMessageSendHandler(event) {
  let externalRecipients = [];

  // Function to check a single email address
  function checkEmail(email, field) {
    // Extract just the email address, handle potential formatting
    let cleanedEmail = email;
    const match = cleanedEmail.match(/<(.+?)>|[^<>\s]+/);
    cleanedEmail = match ? match[1] || match[0] : cleanedEmail;
    cleanedEmail = cleanedEmail.trim().toLowerCase();
    const domain = customerDomain.toLowerCase();
    
    console.log(`Checking ${field} email: ${cleanedEmail}`);
    console.log(`Ends with ${domain}? ${cleanedEmail.endsWith(domain)}`);
    
    if (!cleanedEmail.endsWith(domain)) {
      externalRecipients.push(`${field}: ${cleanedEmail}`);
    }
  }

  // Check "To" recipients
  Office.context.mailbox.item.to.getAsync((toResult) => {
    if (toResult.status === Office.AsyncResultStatus.Succeeded) {
      toResult.value.forEach((recipient) => {
        checkEmail(recipient.emailAddress, "To");
      });

      // Check "CC" recipients
      Office.context.mailbox.item.cc.getAsync((ccResult) => {
        if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
          ccResult.value.forEach((recipient) => {
            checkEmail(recipient.emailAddress, "CC");
          });

          // Check "BCC" recipients
          Office.context.mailbox.item.bcc.getAsync((bccResult) => {
            if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
              bccResult.value.forEach((recipient) => {
                checkEmail(recipient.emailAddress, "BCC");
              });

              // After all checks, decide whether to show popup
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
              console.log("Failed to get BCC recipients, proceeding with checks.");
              // If BCC fails, still check To and CC results
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
            }
          });
        } else {
          console.log("Failed to get CC recipients, proceeding with To check.");
          // If CC fails, still check To results
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
        }
      });
    } else {
      console.log("Failed to get To recipients, allowing send as fallback.");
      event.completed({ allowEvent: true });
    }
  });
}

// Ensure Office API is ready before associating the event handler
Office.onReady((info) => {
  // Check platform and associate the event handler
  if (info.platform === Office.PlatformType.PC || info.platform == null) {
    console.log("Associating onMessageSendHandler for platform: " + (info.platform || "unknown"));
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  } else {
    console.log("Platform not supported for event handler: " + info.platform);
  }
}).catch((error) => {
  console.log("Error initializing Office API: " + error);
});
