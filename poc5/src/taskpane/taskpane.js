Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("apply-not-marketing").onclick = setNotMarketingCustomHeaders;
    document.getElementById("apply-marketing").onclick = setMarketingCustomHeaders;

    Office.context.mailbox.item.to.getAsync(
      { asyncContext: event },
      (asyncResult) => {
        let event = asyncResult.asyncContext;
        let nonGsEmailCount = 0;
        asyncResult.value.forEach((toEmail) => {
          if (toEmail.emailAddress.includes('gs.com') === false) {
            ++nonGsEmailCount;
          }
        });
        if (nonGsEmailCount > 1) {
          document.getElementById("apply-marketing").disabled = true;
          document.getElementById("moreThanOneEmailWarning").style.display = 'block';

        }
      }
    );
  }
});


function setNotMarketingCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "pwm-mar-check": "done", "is-marketing": "no" },
    setCallback
  );
}

function setMarketingCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { 
      "pwm-mar-check": "done", 
      "is-marketing": "yes", 
      "List-Unsubscribe-Post": "List-Unsubscribe=One-Click",
      "List-Unsubscribe": "https://abuthakir4u.github.io/poc5/src/taskpane"
    },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
    //Office.context.mailbox.item.notificationMessages.removeAsync("notificationForMarketingEmail");
    Office.context.mailbox.item.notificationMessages.replaceAsync(
      "notificationForMarketingEmail",
      {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Marketing Acknowledgement done successfully. Please hit Send button now to send email.",
        icon: "icon2",
        persistent: false
      },
      handleResult);
    Office.context.ui.closeContainer();
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

function handleResult(res) {
  console.log("res", res);
}

// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.internetHeaders.getAsync(
    ["pwm-mar-check", "is-marketing"],
    getCallback
  );

  //Todo: Make the below code to work to pass the data to side panel from notification action
  // Office.context.mailbox.item.getInitializationContextAsync((asyncResult) => {
  //   console.log('test');
  //   console.log(asyncResult);

  //   let msg = "log message: " + asyncResult + ", " + asyncResult.value + ", " + JSON.parse(asyncResult.value);

  //   $('#logMsg').html(msg);

  //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //     if (asyncResult.value.length > 0) {
  //       // The value is a string, parse to an object.
  //       console.log('asyncResult', asyncResult);
  //       const context = JSON.parse(asyncResult.value);
  //       console.log('asycontextncResult', context);
  //       // Do something with context.
  //     } else {
  //       // Empty context, treat as no context.
  //     }
  //   } else {
  //     // Handle the error.
  //   }
  // });
}

function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected headers: " + JSON.stringify(asyncResult.value));
    console.log('asyncResult.value', asyncResult.value);
    console.log('asyncResult.value["pwm-mar-check"]', asyncResult.value["pwm-mar-check"])
  } else {
    console.log("Error getting selected headers: " + JSON.stringify(asyncResult.error));
  }
}
