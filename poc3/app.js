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
    console.log('Inside validateToForDNE');
    let dneEntriesInToEmail = [];
    mailboxItem.to.getAsync(
        { asyncContext: event },
        function (asyncResult) {
            if (asyncResult !== null && asyncResult.value.length !== 0) {
                console.log('asyncResult not empty');
                asyncResult.value.forEach((toEmail) => {
                    if (sfDneEmails.indexOf(toEmail.emailAddress) !== -1) {
                        dneEntriesInToEmail.push(toEmail.emailAddress);
                    }
                });
                if (dneEntriesInToEmail.length !== 0) {
                    let commaSepDneEntries = dneEntriesInToEmail.join(', ');
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please remove following DNE emails from recipient list: ' + commaSepDneEntries });
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

