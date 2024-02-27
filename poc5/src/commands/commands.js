Office.onReady();

function onRecipientChangeHandler(event) {
  Office.context.mailbox.item.internetHeaders.setAsync(
    { "pwm-mar-check": "done", "is-marketing": "no" },
    setCallback
  );
}

function setCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Successfully set headers");
    Office.context.mailbox.item.notificationMessages.removeAsync("notificationForMarketingEmail");
  } else {
    console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
  }
}

function onEmailSendHandler(event) {

  Office.context.mailbox.item.internetHeaders.getAsync(
    ["pwm-mar-check"],
    (headerFetchResult) => {
      if (headerFetchResult.status === Office.AsyncResultStatus.Succeeded) {
        if (headerFetchResult.value === null || headerFetchResult.value["pwm-mar-check"] != "done") {
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
              let message = "";
              if (nonGsEmailCount == 0) {
                console.log("No external email found..");
              } else if (nonGsEmailCount == 1) {
                message = 'One external email found. Complete marketing acknowledgement with notification message..';
              } else if (nonGsEmailCount > 1) {
                message = 'More than one external email found. Complete marketing acknowledgement with notification message..';
              }
              if (nonGsEmailCount >= 1) {
                Office.context.mailbox.item.notificationMessages.addAsync("notificationForMarketingEmail", {
                  type: "insightMessage",
                  message: "Please complete marketing email confirmation.",
                  icon: "Icon.16x16",
                  actions: [
                    {
                      actionType: "showTaskPane",
                      actionText: "Acknowledge Margeting",
                      commandId: "MessageComposeSelectButton",
                      contextData: "{'nonGsEmailCount': " +  nonGsEmailCount + "}",
                    }
                  ],
                });
                event.completed({
                  allowEvent: false,
                  errorMessage: message,
                });

              } else {
                console.log("No external email found, so can proceed email send");
                event.completed({
                  allowEvent: true
                });
              }
              return;
            }
          );
        } else {
          console.log("Marketing email header set, so can proceed email send");
          event.completed({
            allowEvent: true
          });
          return;
        }
      } else {
        event.completed({
          allowEvent: false,
          errorMessage: "Unable to read headers...",
        });
        return;
      }
    }
  );
}

//Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
//Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
//Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
Office.actions.associate("onMessageSendHandler", onEmailSendHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onRecipientChangeHandler);
