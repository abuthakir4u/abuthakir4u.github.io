/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let mailboxItem;

const sfDneEmails = ['abu@gmail.com', 'chhavi@gmail.com'];

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

function validateToForDNE(event) {
    let dneEntriesInToEmail = [];
    mailboxItem.to.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            console.log('asyncResult', asyncResult);
            if (asyncResult !== null && asyncResult.value.length !== 0) {
                asyncResult.value.forEach((toEmail) => {
                    if (sfDneEmails.indexOf(toEmail) != -1) {
                        dneEntriesInToEmail.push(toEmail);
                    }
                });
                if (dneEntriesInToEmail.length !== 0) {
                    console.log('DNE entries found');
                    let commaSepDneEntries = dneEntriesInToEmail.join(', ');
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please remove following DNE emails from To: ' + commaSepDneEntries });
                    asyncResult.asyncContext.completed({ allowEvent: false });
                } else {
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }
            } else {
                console.log("No DNE entry found, so can proceed");
                asyncResult.asyncContext.completed({ allowEvent: true });
            }
        }
    );
}

