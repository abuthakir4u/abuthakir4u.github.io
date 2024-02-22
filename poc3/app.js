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
            console.log('asyncResult', asyncResult);
            console.log('asyncResult.value', asyncResult.value);
            console.log('asyncResult.value[0].email', asyncResult.value[0].email);
            if (asyncResult !== null && asyncResult.value.length !== 0) {
                console.log('asyncResult not empty');
                asyncResult.value.forEach((toEmail) => {
                    console.log('looping to find entry of '+ toEmail);
                    if (sfDneEmails.indexOf(toEmail.emailAddress) !== -1) {
                        dneEntriesInToEmail.push(toEmail);
                    }
                });
                console.log('dneEntriesInToEmail', dneEntriesInToEmail);
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
                asyncResult.asyncContext.completed({ allowEvent: trfalseue });
            }
        }
    );
}

