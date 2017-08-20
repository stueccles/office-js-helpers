export declare class UI {
    /** Shows a basic notification at the top of the page
     * @param message - Message, either single-string or multiline (punctuated by '\n')
     */
    static notify(message: string): any;
    /** Shows a basic error notification at the top of the page
     * @param error - Error object
     */
    static notify(error: Error): any;
    /** Shows a basic notification with a custom title at the top of the page
     * @param title - Title, bolded
     * @param message - Message, either single-string or multiline (punctuated by '\n')
    */
    static notify(title: string, message: string): any;
    /** Shows a basic error notification with a custom title at the top of the page
     * @param title - Title, bolded
     * @param error - Error object
     */
    static notify(title: string, error: Error): any;
    /** Shows a basic error notification, with custom parameters, at the top of the page */
    static notify(error: Error, params: {
        title?: string;
        /** custom message in place of the error text */
        message?: string;
        moreDetailsLabel?: string;
    }): any;
    /** Shows a basic notification at the top of the page, with a background color set based on the type parameter
     * @param title - Title, bolded
     * @param message - Message, either single-string or multiline (punctuated by '\n')
     * @param type - Type, determines the background color of the notification. Acceptable types are:
     *               'default' | 'success' | 'error' | 'warning' | 'severe-warning'
     */
    static notify(title: string, message: string, type: 'default' | 'success' | 'error' | 'warning' | 'severe-warning'): any;
    /** Shows a basic notification at the top of the page, with custom parameters */
    static notify(params: {
        title?: string;
        message: string;
        type?: 'default' | 'success' | 'error' | 'warning' | 'severe-warning';
    }): any;
}
