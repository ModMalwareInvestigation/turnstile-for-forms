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

