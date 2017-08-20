import { Storage } from '../helpers/storage';
export declare const DefaultEndpoints: {
    Google: string;
    Microsoft: string;
    Facebook: string;
    AzureAD: string;
};
export interface IEndpointConfiguration {
    /**
     * Unique name for the Endpoint
     */
    provider?: string;
    /**
     * Registered OAuth ClientID
     */
    clientId?: string;
    /**
     * Base URL of the endpoint
     */
    baseUrl?: string;
    /**
     * URL segment for OAuth authorize endpoint.
     * The final authorize url is constructed as (baseUrl + '/' + authorizeUrl).
     */
    authorizeUrl?: string;
    /**
     * Registered OAuth redirect url.
     * Defaults to window.location.origin
     */
    redirectUrl?: string;
    /**
     * Optional token url to exchange a code with.
     * Not recommended if OAuth provider supports implicit flow.
     */
    tokenUrl?: string;
    /**
     * Registered OAuth scope.
     */
    scope?: string;
    /**
     * Resource paramater for the OAuth provider.
     */
    resource?: string;
    /**
     * Automatically generate a state? defaults to false.
     */
    state?: boolean;
    /**
     * Automatically generate a nonce? defaults to false.
     */
    nonce?: boolean;
    /**
     * OAuth responseType.
     */
    responseType?: string;
    /**
     * OAuth grantType.
     */
    grantType?: string;
    /**
     * Additional object for query parameters.
     * Will be appending them after encoding the values.
     */
    extraQueryParameters?: {
        [index: string]: string;
    };
    /**
     * Additional object for headers.
     */
    extraHeaders?: {
        [index: string]: string;
    };
}
/**
 * Helper for creating and registering OAuth Endpoints.
 */
export declare class EndpointStorage extends Storage<IEndpointConfiguration> {
    /**
     * @constructor
    */
    constructor();
    /**
     * Extends Storage's default add method.
     * Registers a new OAuth Endpoint.
     *
     * @param {string} provider Unique name for the registered OAuth Endpoint.
     * @param {object} config Valid Endpoint configuration.
     * @see {@link IEndpointConfiguration}.
     * @return {object} Returns the added endpoint.
     */
    add(provider: string, config: IEndpointConfiguration): IEndpointConfiguration;
    /**
     * Register Google Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Google App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    registerGoogleAuth(clientId: string, overrides?: IEndpointConfiguration): IEndpointConfiguration;
    /**
     * Register Microsoft Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Microsoft App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    registerMicrosoftAuth(clientId: string, overrides?: IEndpointConfiguration): void;
    /**
     * Register Facebook Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Facebook App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    registerFacebookAuth(clientId: string, overrides?: IEndpointConfiguration): void;
    /**
     * Register AzureAD Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the AzureAD App.
     * @param {string} tenant Tenant for the AzureAD App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    registerAzureADAuth(clientId: string, tenant: string, overrides?: IEndpointConfiguration): void;
    /**
     * Helper to generate the OAuth login url.
     *
     * @param {object} config Valid Endpoint configuration.
     * @return {object} Returns the added endpoint.
     */
    static getLoginParams(endpointConfig: IEndpointConfiguration): {
        url: string;
        state: number;
    };
    static generateCryptoSafeRandom(): number;
}
