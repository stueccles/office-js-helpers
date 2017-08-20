import { EndpointStorage } from './endpoint.manager';
import { TokenStorage, IToken, ICode, IError } from './token.manager';
import { CustomError } from '../errors/custom.error';
/**
 * Custom error type to handle OAuth specific errors.
 */
export declare class AuthError extends CustomError {
    innerError: Error;
    /**
     * @constructor
     *
     * @param message Error message to be propagated.
     * @param state OAuth state if available.
    */
    constructor(message: string, innerError?: Error);
}
/**
 * Helper for performing Implicit OAuth Authentication with registered endpoints.
 */
export declare class Authenticator {
    endpoints: EndpointStorage;
    tokens: TokenStorage;
    /**
     * @constructor
     *
     * @param endpoints Depends on an instance of EndpointStorage.
     * @param tokens Depends on an instance of TokenStorage.
    */
    constructor(endpoints?: EndpointStorage, tokens?: TokenStorage);
    /**
     * Authenticate based on the given provider.
     * Either uses DialogAPI or Window Popups based on where its being called from either Add-in or Web.
     * If the token was cached, the it retrieves the cached token.
     * If the cached token has expired then the authentication dialog is displayed.
     *
     * NOTE: you have to manually check the expires_in or expires_at property to determine
     * if the token has expired.
     *
     * @param {string} provider Link to the provider.
     * @param {boolean} force Force re-authentication.
     * @return {Promise<IToken|ICode>} Returns a promise of the token or code or error.
     */
    authenticate(provider: string, force?: boolean, useMicrosoftTeams?: boolean): Promise<IToken>;
    /**
     * Check if the currrent url is running inside of a Dialog that contains an access_token or code or error.
     * If true then it calls messageParent by extracting the token information, thereby closing the dialog.
     * Otherwise, the caller should proceed with normal initialization of their application.
     *
     * @return {boolean}
     * Returns false if the code is running inside of a dialog without the required information
     * or is not running inside of a dialog at all.
     */
    static isAuthDialog(useMicrosoftTeams?: boolean): boolean;
    /**
     * Extract the token from the URL
     *
     * @param {string} url The url to extract the token from.
     * @param {string} exclude Exclude a particlaur string from the url, such as a query param or specific substring.
     * @param {string} delimiter[optional] Delimiter used by OAuth provider to mark the beginning of token response. Defaults to #.
     * @return {object} Returns the extracted token.
     */
    static getUrlParams(url?: string, exclude?: string, delimiter?: string): ICode | IToken | IError;
    static extractParams(segment: string): any;
    private _openAuthDialog(provider, useMicrosoftTeams);
    private _openInWindowPopup(provider);
    /**
     * Helper for exchanging the code with a registered Endpoint.
     * The helper sends a POST request to the given Endpoint's tokenUrl.
     *
     * The Endpoint must accept the data JSON input and return an 'access_token'
     * in the JSON output.
     *
     * @param {Endpoint} endpoint Endpoint configuration.
     * @param {object} data Data to be sent to the tokenUrl.
     * @param {object} headers Headers to be sent to the tokenUrl.     *
     * @return {Promise<IToken>} Returns a promise of the token or error.
     */
    private _exchangeCodeForToken(endpoint, data, headers?);
    private _handleTokenResult(redirectUrl, endpoint, state);
}
