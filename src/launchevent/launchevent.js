// Define the customer domain
const customerDomain = "@bizwind.co.jp";

// Utility to add a timeout to a promise
const withTimeout = (promise, ms) => {
  const timeout = new Promise((_, reject) =>
    setTimeout(() => reject(new Error("Operation timed out")), ms)
  );
  return Promise.race([promise, timeout]);
};

// Handler for the OnMessageSend event
async function onMessageSendHandler(event) {
  const handlerTimeout = 4000; // 4 seconds to stay within Outlook's ~5s limit
  try {
    // Wrap the entire operation in a timeout
    const result = await withTimeout(
      (async () => {
        let externalRecipients = [];
        let externalDomains = new Set();

        // Function to check a single email address
        function checkEmail(email, field) {
          let cleanedEmail = email.trim().toLowerCase();
          // Robust email extraction
          const match = cleanedEmail.match(/<(.+?)>|([^\s<>]+)/);
          cleanedEmail = match ? (match[1] || match[0]) : cleanedEmail;

          // Basic email validation
          if (!cleanedEmail.includes("@")) {
            console.log(`Invalid email skipped: ${cleanedEmail}`);
            return;
          }

          const domain = customerDomain.toLowerCase();
          console.log(`Checking ${field} email: ${cleanedEmail}`);
          if (!cleanedEmail.endsWith(domain)) {
            externalRecipients.push(`${field}: ${cleanedEmail}`);
            const emailDomain = `@${cleanedEmail.split('@')[1]}`;
            externalDomains.add(emailDomain);
          }
        }

        // Parallelize recipient retrieval with individual timeouts
        const [toResult, ccResult, bccResult] = await Promise.all([
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.to.getAsync(resolve)),
            1500
          ),
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.cc.getAsync(resolve)),
            1500
          ),
          withTimeout(
            new Promise((resolve) => Office.context.mailbox.item.bcc.getAsync(resolve)),
            1500
          ),
        ]);

        // Process recipients
        if (toResult.status === Office.AsyncResultStatus.Succeeded) {
          toResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "To"));
        } else {
          console.log("Failed to get To recipients:", toResult.error?.message);
        }

        if (ccResult.status === Office.AsyncResultStatus.Succeeded) {
          ccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "CC"));
        } else {
          console.log("Failed to get CC recipients:", ccResult.error?.message);
        }

        if (bccResult.status === Office.AsyncResultStatus.Succeeded) {
          bccResult.value.forEach((recipient) => checkEmail(recipient.emailAddress, "BCC"));
        } else {
          console.log("Failed to get BCC recipients:", bccResult.error?.message);
        }

        // Decide whether to show popup
        if (externalRecipients.length > 0) {
          const message =
            "You are sending this email to external recipients:\n\n" +
            "Domain list\n" +
            Array.from(externalDomains).join("\n") +
            "\n\nEmail list\n" +
            externalRecipients.join("\n") +
            "\n\nAre you sure you want to send it?";
          return { allowEvent: false, errorMessage: message };
        } else {
          console.log("No external recipients found, allowing send.");
          return { allowEvent: true };
        }
      })(),
      handlerTimeout
    );

    event.completed(result);
  } catch (error) {
    console.error("Error in onMessageSendHandler:", error.message, error.stack);
    // Block send on error to prevent accidental external sends
    event.completed({
      allowEvent: false,
      errorMessage: "Error: Unable to verify recipients. Please try again or contact support.",
    });
  }
}

// Initialize Office API
Office.onReady((info) => {
  if (info.platform === Office.PlatformType.PC || info.platform === null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
}).catch((error) => {
  console.error("Error initializing Office API:", error.message, error.stack);
});
