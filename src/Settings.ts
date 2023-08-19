type Settings = {
    /**
     * Enable the add-on for the current spreadsheet
     * Per-document property
     */
    enableTurnstile: boolean
    /**
     * Marks the current spreadsheet as the place where turnstile tokens will be stored
     * The add-on supports only one token sheet.
     * Per-document property
     */
    isTokenSheet: boolean
    /**
     * The Cloudflare Turnstile secret key
     * See https://developers.cloudflare.com/turnstile/get-started/#get-a-sitekey-and-secret-key
     */
    siteSecret: string
    /**
     * Enable Discord webhook notifications
     * Turnstile-validated responses will be sent to the webhooks specified in {@code webhookUrls}
     */
    enableNotifications: boolean
    /**
     * A newline-separated list of Discord webhook URLs to send turnstile-validated form responses to
     * Per-document property
     */
    webhookUrls: string
    /**
     * A newline-separated list of form questions that determine the order in which form responses appear on the Discord
     * webhook notification
     */
    responseOrder: string
    /**
     * The name of the notification template to use for Discord webhook notifications
     * @see WebhookPayloadTemplates.ts
     */
    notificationTemplate: string
}