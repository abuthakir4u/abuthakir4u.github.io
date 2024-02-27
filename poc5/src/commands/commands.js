/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady();

/**
 * The words in the subject or body that require corresponding color categories to be applied to a new
 * message or appointment.
 * @constant
 * @type {string[]}
 */
const KEYWORDS = [
  "sales",
  "expense reports",
  "legal",
  "marketing",
  "performance reviews",
];

/**
 * Handle the OnNewMessageCompose or OnNewAppointmentOrganizer event by verifying that keywords have corresponding
 * color categories when a new message or appointment is created. If no corresponding categories exist, they will be
 * created.
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
 */
function onItemComposeHandler(event) {
  Office.context.mailbox.masterCategories.getAsync(
    { asyncContext: event },
    (asyncResult) => {
      let event = asyncResult.asyncContext;

      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        event.completed({
          allowEvent: false,
          errorMessage: "Failed to configure categories.",
        });
        return;
      }

      let categories = asyncResult.value;
      let categoriesToBeCreated = [];
      if (categories) {
        let categoryNamesInUse = getCategoryProperty(categories, "displayName");
        let categoryColorsInUse = getCategoryProperty(categories, "color");
        categoriesToBeCreated = getCategoriesToBeCreated(
          KEYWORDS,
          categoryNamesInUse
        );

        if (categoriesToBeCreated.length > 0) {
          categoriesToBeCreated = assignCategoryColors(
            categoriesToBeCreated,
            categoryColorsInUse
          );
        }
      } else {
        categoriesToBeCreated = assignCategoryColors(
          getCategoriesToBeCreated(KEYWORDS)
        );
      }

      createCategories(event, categoriesToBeCreated);
      event.completed({ allowEvent: true });
    }
  );
}

/**
 * Handle the OnMessageSend or OnAppointmentSend event by verifying that applicable color categories are
 * applied to a new message or appointment before it's sent.
 * @param {Office.AddinCommands.Event} event The OnMessageSend or OnAppointmentSend event object.
 */

function onItemSendHandler(event) {

  Office.context.mailbox.item.internetHeaders.getAsync(
    ["pwm-mar-check"],
    (headerFetchResult) => {
      if (headerFetchResult.status === Office.AsyncResultStatus.Succeeded) {
        if (headerFetchResult.value != null && headerFetchResult.value["pwm-mar-check"] != "done") {
          Office.context.mailbox.item.to.getAsync(
            { asyncContext: event },
            (asyncResult) => {
              let event = asyncResult.asyncContext;
              let nonGsEmailCount = 0;
              asyncResult.value.forEach((toEmail) => {
                if (toEmail.emailAddress.includes('@gs.com') === false) {
                  ++nonGsEmailCount;
                }
              });

              if (nonGsEmailCount == 0) {
                message = 'This is internal email, so no need to do anything..';
              } else if (nonGsEmailCount == 1) {
                message = 'One external email found. Completed marketing acknowledgement with notification message..';
              } else if (nonGsEmailCount > 1) {
                message = 'More than one external email found. Completed marketing acknowledgement with notification message..';
              }

              Office.context.mailbox.item.notificationMessages.addAsync("notificationForMarketingEmail", {
                type: "insightMessage",
                message: "Please complete marketing email confirmation.",
                icon: "Icon.16x16",
                actions: [
                  {
                    actionType: "showTaskPane",
                    actionText: "Acknowledge Margeting",
                    commandId: "MessageComposeSelectButton",
                    contextData: "{''}",
                  },
                ],
              });

              event.completed({
                allowEvent: false,
                errorMessage: message,
              });
              return;
            }
          );
        } else {
          event.completed({
            allowEvent: false,
            errorMessage: "Marketing email header set, so can proceed email sending",
          });
          return;
        }
      } else {
        //Todo: handle error
      }
    }
  );
}

//Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
//Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
Office.actions.associate("onMessageSendHandler", onItemSendHandler);
//Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
