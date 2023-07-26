/**
 * Sends a notification to all the webhooks stored in the webhookUrls document property
 *
 * @param e the event containing the form response to send
 */
function sendNotification(e: GoogleAppsScript.Events.SheetsOnFormSubmit) {
    let webhooks = PropertiesService.getDocumentProperties().getProperty('webhookUrls');
    let templateName = PropertiesService.getDocumentProperties().getProperty('notificationTemplate');
    let responseOrder = PropertiesService.getDocumentProperties().getProperty('responseOrder');
    let substitution = {
        formTitle: PropertiesService.getDocumentProperties().getProperty('formTitle'),
        sessionId: e.namedValues['Secret'][0],
        questionResponses: '',
        // https://stackoverflow.com/questions/64182411/current-timestamp-in-iso-8601-utc-time-ie-2013-11-03t0045540200
        timestamp: new Date(Utilities.formatDate(new Date(e.namedValues['Timestamp'][0]), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX")).toISOString()
    }

    delete e.namedValues['Secret'];
    delete e.namedValues['Timestamp'];

    if (webhooks === null) {
        throw new Error('Failed to send notifications, no webhooks were provided');
    } else if (templateName === null) {
        throw new Error('Failed to send notifications, no template name provided');
    } else if (responseOrder === null) {
        throw new Error('Failed to send notifications, no response order provided');
    }

    let webhookList = webhooks.split('\n');
    let responseOrderList = responseOrder.split('\n');

    for (const questionName of responseOrderList) {
        const response = e.namedValues[questionName];

        if (typeof response !== 'undefined') {
            let responseValue = '';

            for (const responseElement of response) {
                if (responseElement !== '') {
                    responseValue = `${responseValue}${responseElement}\n`
                }
            }

            if (responseValue.length > 1) {
                responseValue = responseValue.substring(0, responseValue.length - 1);
            }

            if (responseValue !== '') {
                substitution.questionResponses = `${substitution.questionResponses}{"name": "${questionName}", "value": "${responseValue}"},`;
            }
        } else {
            console.warn(`Question "${questionName}" in document property "responseOrder" was not found.`);
        }
    }

    // Remove the trailing comma
    if (substitution.questionResponses.length > 0) {
        substitution.questionResponses = substitution.questionResponses.substring(0, substitution.questionResponses.length - 1);
    } else {
        // Add a default message
        substitution.questionResponses = '{"name": "No Answers Provided", "value": "User did not fill in any questions"}';
    }

    for (const webhook of webhookList) {
        if (webhook.indexOf('https://discord.com/api/webhooks') === 0) {
            //@ts-ignore
            postToWebhook(webhook, createWebhookPayload(templateName, substitution));
        } else {
            throw new Error(`Failed to post to url '${webhook}'. URL is not a valid Discord webhook.`);
        }
    }
}

/**
 * Substitutes the contents of the given object into a webhook payload template to create a complete payload
 *
 * @param templateName name of the template to use
 * @param substitution object to substitute into the template
 */
function createWebhookPayload(templateName: string, substitution: object): string {
    return evaluatePayloadTemplate(templateName, substitution)
}

function postToWebhook(url: string, payload: string) {
    let options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        'method' : 'post',
        'contentType': 'application/json',
        'payload' : payload,
        'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode < 200 || responseCode >= 300) {
        throw new Error(`Post to webhook failed with code ${response.getResponseCode()}\n${response.getContentText()}\nFailing Request Payload:\n${payload}`);
    }
}

