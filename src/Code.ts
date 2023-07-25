/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
const ADDON_TITLE = 'Turnstile for Forms';
const NOTIFICATION_TEMPLATES = {'Report Form': 'reportFormNotification'};

/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen | undefined) {
    SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Configure', 'showSidebar')
        .addItem('About', 'showAbout')
        .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e: GoogleAppsScript.Events.AddonOnInstall) {
    onOpen(undefined);
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
    const ui = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle(ADDON_TITLE);
    SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
    const ui = HtmlService.createHtmlOutputFromFile('about')
        .setWidth(420)
        .setHeight(270);
    SpreadsheetApp.getUi().showModalDialog(ui, `About ${ADDON_TITLE}`);
}

/**
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(settings: any) {
    const sheetTitle = SpreadsheetApp.getActive().getName();
    const formTitle = sheetTitle.substring(0, sheetTitle.indexOf('(')).trim();

    PropertiesService.getDocumentProperties().setProperties({
        'enableTurnstile': settings.enableTurnstile,
        'isTokenSheet': settings.isTokenSheet,
        'siteSecret': settings.siteSecret
    });

    if (settings.isTokenSheet) {
        if (SpreadsheetApp.getActive().getFormUrl() !== null) {
            PropertiesService.getUserProperties().setProperty('tokenFormUrl', SpreadsheetApp.getActive().getFormUrl() as string);
            console.log(SpreadsheetApp.getActive().getFormUrl());
        }
        else {
            throw new Error('Token sheet is not associated with a token form');
        }
    }
    else {
        if (settings.enableTurnstile) {
            if (settings.enableNotifications) {
                PropertiesService.getDocumentProperties().setProperty('enableNotifications', settings.enableNotifications);

                if (!Object.values(NOTIFICATION_TEMPLATES).includes(settings.notificationTemplate)) {
                    throw new Error(`Unknown notification template ${settings.notificationTemplate}`);
                }

                PropertiesService.getDocumentProperties().setProperties({
                    'webhookUrls': settings.webhookUrls,
                    'formTitle': formTitle,
                    'notificationTemplate': settings.notificationTemplate
                });
            }
        }
        else {
            PropertiesService.getDocumentProperties().setProperty('enableNotifications', 'false');
        }
    }

    adjustClearTokenFormTrigger(settings.isTokenSheet);
    adjustFormSubmitTrigger();
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
    const settings: any = PropertiesService.getDocumentProperties().getProperties();
    settings.notificationTemplates = NOTIFICATION_TEMPLATES;

    return settings;
}

function getSiteSecret() {
    try {
        return PropertiesService.getUserProperties().getProperty('siteSecret');
    } catch (e) {
        throw new Error(`Load site secret failed: \n${(e as Error).stack}`)
    }
}

/**
 * Adjust the onFormSubmit triggers for the form and token form based on user's requests.
 */
function adjustFormSubmitTrigger() {
    const sheet = SpreadsheetApp.getActive();
    const triggers = ScriptApp.getUserTriggers(sheet);
    const settings = PropertiesService.getDocumentProperties();
    const triggerNeeded =
        settings.getProperty('enableTurnstile') === 'true';

    // Create a new trigger if required; delete existing trigger
    // if it is not needed.
    let existingTrigger = null;
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT) {
            existingTrigger = triggers[i];
            break;
        }
    }
    if (triggerNeeded && !existingTrigger) {
        ScriptApp.newTrigger('respondToFormSubmit')
            .forSpreadsheet(sheet)
            .onFormSubmit()
            .create();
    } else if (!triggerNeeded && existingTrigger) {
        ScriptApp.deleteTrigger(existingTrigger);
    }
}

/**
 * If the active form is the token form, set a trigger to clear it once a day.
 */
function adjustClearTokenFormTrigger(isTokenForm: boolean) {
    const sheet = SpreadsheetApp.getActive();
    const triggers = ScriptApp.getUserTriggers(sheet);

    // Create a new trigger if required; delete existing trigger
    // if it is not needed.
    let existingTrigger = null;
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
            existingTrigger = triggers[i];
            break;
        }
    }
    if (isTokenForm && !existingTrigger) {
        ScriptApp.newTrigger('setClearTokenFormFlag')
            .timeBased()
            .everyDays(1)
            .create();
    } else if (!isTokenForm && existingTrigger) {
        ScriptApp.deleteTrigger(existingTrigger);
    }
}

/**
 * Clear the token form's responses except the newest 20 responses
 */
function clearTokenFormResponses() {
    const tokenSheet = SpreadsheetApp.getActive();
    const tokenForm = FormApp.openByUrl(tokenSheet.getFormUrl() as string);
    const responseSheet = tokenSheet.getActiveSheet();
    const numberOfResponses = tokenForm.getResponses().length;
    let numberToDelete = numberOfResponses - 20;

    if (numberToDelete > 0) {
        console.info(`Will delete ${numberToDelete}/${numberOfResponses} tokens`);
        for (let i = 0; i < numberToDelete; i++) {

            console.log(`Deleting form response ${i}`);
            tokenForm.deleteResponse(tokenForm.getResponses()[i].getId());
        }
        responseSheet.deleteRows(2, numberToDelete);
        console.info('Cleared token form responses');
    } else {
        console.info('Skipped clearing token form responses: too few responses')
    }
}

/**
 * Responds to a form submission event if an onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToFormSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit) {
    const settings = PropertiesService.getDocumentProperties();
    const tokenFormUrl = PropertiesService.getUserProperties().getProperty('tokenFormUrl');
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

    // Check if the actions of the trigger require authorizations that have not
    // been supplied yet -- if so, warn the active user via email (if possible).
    // This check is required when using triggers with add-ons to maintain
    // functional triggers.
    if (authInfo.getAuthorizationStatus() ===
        ScriptApp.AuthorizationStatus.REQUIRED) {
        // Re-authorization is required. In this case, the user needs to be alerted
        // that they need to reauthorize; the normal trigger action is not
        // conducted, since authorization needs to be provided first. Send at
        // most one 'Authorization Required' email a day, to avoid spamming users
        // of the add-on.
        sendReauthorizationRequest();
    } else {
        // All required authorizations have been granted, so continue to respond to
        // the trigger event.
        if (typeof tokenFormUrl === null) {
            throw new Error('Token sheet is not associated with a token form');
        }

        const tokenForm = FormApp.openByUrl(tokenFormUrl as string);

        // Make sure this isn't the token form
        if (settings.getProperty('isTokenSheet') !== 'true') {
            if (settings.getProperty('enableTurnstile') === 'true') {
                const uuid = String(e.namedValues['Secret']);
                const responseStartCoordinate = e.range.getCell(1, 1).getA1Notation();
                const errorBackgroundColor = '#f4cccc';

                if (typeof uuid !== 'undefined') {
                    const token = getLatestToken(tokenForm, uuid);

                    if (token) {
                        if (!verifyTurnstileToken(uuid, token)) {
                            e.range.setBackground(errorBackgroundColor);
                            console.warn(`Response at ${responseStartCoordinate} failed verification`);
                        } else {
                            console.info(`Response at ${responseStartCoordinate} passed verification`);
                            copyResponseToFilteredResponseSheet(e.range);
                            sendNotification(e);
                        }
                    } else {
                        e.range.setBackground(errorBackgroundColor);
                        console.warn(`Failed to find token for response with UUID "${uuid}"`)
                    }
                } else {
                    e.range.setBackground(errorBackgroundColor);
                    throw new Error(`Response at ${responseStartCoordinate} didn't include a UUID.`);
                }
            }
        } else {
            if (e.namedValues['UUID']) {
                console.log(`Received token for request ${e.namedValues['UUID']}`);

                if (PropertiesService.getUserProperties().getProperty('clearTokenFormOnNextRun') === 'true') {
                    console.info('Running scheduled token cleanup...');
                    clearTokenFormResponses();
                    PropertiesService.getUserProperties().setProperty('clearTokenFormOnNextRun', 'false');
                } else if (tokenForm.getResponses().length > 20000) {
                    console.info('Token count exceeded 20K, running cleanup...');
                    clearTokenFormResponses();
                    PropertiesService.getUserProperties().setProperty('clearTokenFormOnNextRun', 'false');
                }
            } else {
                throw new Error('Token form response has no UUID');
            }
        }
    }
}

function sendNotification(e: GoogleAppsScript.Events.SheetsOnFormSubmit) {
    const webhooks = PropertiesService.getDocumentProperties().getProperty('webhookUrls')!.split(',');
    const templateName = PropertiesService.getDocumentProperties().getProperty('notificationTemplate');
    let substitution = {
        formTitle: PropertiesService.getDocumentProperties().getProperty('formTitle'),
        sessionId: e.namedValues['Secret'][0],
        questionResponses: '',
        // Hardcoded time zone isn't ideal but time zone conversion requires libraries which is less ideal.
        // Get form TZ: SpreadsheetApp.getActive().getSpreadsheetTimeZone()
        timestamp: new Date(e.namedValues['Timestamp'][0] + ' EDT').toISOString()
    }

    delete e.namedValues['Secret'];
    delete e.namedValues['Timestamp'];

    for (const [key, value] of Object.entries(e.namedValues)) {
        substitution.questionResponses = `${substitution.questionResponses}{"name": "${key}", "value": "${value[0]}"},`;
    }

    // Remove the trailing comma
    if (substitution.questionResponses.length > 0) {
        substitution.questionResponses = substitution.questionResponses.substring(0, substitution.questionResponses.length - 1);
    } else {
        // Add a default message
        substitution.questionResponses = '{"name": "No Answers Provided", "value": "User did not fill in any questions"}';
    }

    for (const webhook of webhooks) {
        if (webhook.indexOf('https://discord.com/api/webhooks') === 0) {
            //@ts-ignore
            postToWebhook(webhook, createWebhookPayload(templateName, substitution));
        } else {
            throw new Error(`Failed to post to url '${webhook}'. URL is not a valid Discord webhook.`);
        }
    }
}


/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
function sendReauthorizationRequest() {
    const settings = PropertiesService.getDocumentProperties();
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    const lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
    const today = new Date().toDateString();
    if (lastAuthEmailDate !== today) {
        if (MailApp.getRemainingDailyQuota() > 0) {
            const template =
                HtmlService.createTemplateFromFile('authorizationEmail');
            template.url = authInfo.getAuthorizationUrl();
            const message = template.evaluate();
            MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
                'Authorization Required',
                message.getContent(), {
                    name: ADDON_TITLE,
                    htmlBody: message.getContent()
                });
        }
        settings.setProperty('lastAuthEmailDate', today);
    }
}

/**
 * Copies the response within the given range to the sheet for filtered (Turnstile verified) responses.
 * The responses are copied to a sheet named "Filtered Responses," which is created if it does not exist.
 */
function copyResponseToFilteredResponseSheet(range: GoogleAppsScript.Spreadsheet.Range) {
    let filteredSheet = SpreadsheetApp.getActive().getSheetByName('Filtered Responses');
    let rangeWidth = range.getWidth();

    if (!filteredSheet) {
        let header = SpreadsheetApp.getActiveSheet().getRange(`A1:A${rangeWidth}`).getValues();
        filteredSheet = SpreadsheetApp.getActive().insertSheet('Filtered Responses');
        // Copy the table header
        filteredSheet.getRange(`A1:A${rangeWidth}`).setValues(header);
        console.log('Created filtered responses sheet');
    }

    filteredSheet.appendRow(range.getValues()[0]);
    // @ts-ignore
    console.log(`Inserted entry at ${filteredSheet.getActiveRange().getA1Notation()} in filtered response sheet`);
}

/**
 * Set a flag to clear the token form on the next token submission.
 * The form is not cleared directly since that would require an OAuth scope to access all sheets on the account.
 */
function setClearTokenFormFlag(e: GoogleAppsScript.Script.EventType.CLOCK) {
    const flag = PropertiesService.getUserProperties().getProperty('clearTokenFormOnNextRun');

    if (flag !== 'true') {
        PropertiesService.getUserProperties().setProperty('clearTokenFormOnNextRun', 'true');
    }
}

/**
 * Verifies a Cloudflare Turnstile token against the siteverify API
 */
function verifyTurnstileToken(uuid: string, token: string) {
    const turnstileSiteVerifyUrl = 'https://challenges.cloudflare.com/turnstile/v0/siteverify';
    const secret = getSiteSecret();

    if (!token) {
        throw new Error('Token verification failed: token is undefined');
    }

    const formData = {
        'secret': secret,
        'response': token
    }

    const options = {
        method: 'post' as GoogleAppsScript.URL_Fetch.HttpMethod,
        payload: formData,
        muteHttpExceptions: true
    }

    let response = UrlFetchApp.fetch(turnstileSiteVerifyUrl, options);
    let responseContent = JSON.parse(response.getContentText());

    if (responseContent['success']) {
        if (responseContent['cdata'] === uuid) {
            return true;
        } else {
            throw new Error('Session ID mismatch between form and Cloudflare');
        }
    } else {
        let errorMessage = `CF replied with '${responseContent['error-codes']}'`;

        if (responseContent['messages'].length > 0) {
            for (let cfErrorMessage of responseContent['messages']) {
                errorMessage = `${errorMessage}\n${cfErrorMessage}`
            }
        }

        console.warn(errorMessage);
    }

    return false;
}

/**
 * Get the token from the most recent response containing the given UUID in the given form
 */
function getLatestToken(form: GoogleAppsScript.Forms.Form, uuid: string) {
    if (!form) {
        throw new Error('Failed to get token: form is null');
    } else if (!uuid) {
        throw new Error('Failed to get token: UUID is null');
    }

    const responses = form.getResponses();

    // Entries are sorted from earliest to latest.
    for (let i = form.getResponses().length - 1; i >= 0; i--) {
        const response = responses[i];
        const itemResponses = response.getItemResponses();
        let uuidResponse: GoogleAppsScript.Forms.ItemResponse | null = null;
        let tokenResponse: GoogleAppsScript.Forms.ItemResponse | null = null;

        if (itemResponses.length !== 2) {
            throw new Error(`Expected token form to have two fields, received ${itemResponses.length} fields`);
        }

        for (const itemResponse of itemResponses) {
            if (itemResponse.getItem().getTitle() === 'UUID') {
                uuidResponse = itemResponse;
            } else if (itemResponse.getItem().getTitle() === 'Token') {
                tokenResponse = itemResponse;
            }
        }

        if (uuidResponse != null && uuidResponse.getResponse() === uuid) {
            if (tokenResponse != null && tokenResponse.getResponse() != null) {
                return tokenResponse.getResponse() as string;
            }
        }
    }

    return null;
}