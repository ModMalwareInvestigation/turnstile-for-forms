let webhookPayloadTemplates = new Map<string, string>();

// Templates
webhookPayloadTemplates.set('reportFormNotification', '{"username": "Notifications for Forms","embeds": [{"title": "${formTitle}","fields": [${questionResponses}],"footer": {"text": "${sessionId}"},"timestamp": "${timestamp}"}]}');

/**
 * Replaces placeholders in a string template with the values of an object that has keys matching the placeholder names in the substitution object
 *
 * @param templateName name of the template
 * @param substitution object with values to substitute into the template
 */
function evaluatePayloadTemplate(templateName: string, substitution: object): string {
    let template = webhookPayloadTemplates.get(templateName);

    if (template) {
        for (const [key, value] of Object.entries(substitution)) {
            template = template.replace(new RegExp(`\\\${${key}}`), value);
        }
    } else {
        throw new Error(`Unknown template "${templateName}"`);
    }

    return template;
}