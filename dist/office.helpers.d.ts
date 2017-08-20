declare module '@microsoft/office-js-helpers' {
  export = OfficeHelpers;
}
declare namespace OfficeHelpers {
  /**
   * Custom error type to handle OAuth specific errors.
   */
  export class AuthError extends CustomError {
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
  export class Authenticator {
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
  export const DefaultEndpoints: {
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
  export class EndpointStorage extends Storage<IEndpointConfiguration> {
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
  export interface IToken {
      provider: string;
      id_token?: string;
      access_token?: string;
      token_type?: string;
      scope?: string;
      state?: string;
      expires_in?: string;
      expires_at?: Date;
  }
  export interface ICode {
      provider: string;
      code: string;
      scope?: string;
      state?: string;
      grantType?: string;
  }
  export interface IError {
      error: string;
      state?: string;
  }
  /**
   * Helper for caching and managing OAuth Tokens.
   */
  export class TokenStorage extends Storage<IToken> {
      /**
       * @constructor
      */
      constructor();
      /**
       * Compute the expiration date based on the expires_in field in a OAuth token.
       */
      static setExpiry(token: IToken): void;
      /**
       * Check if an OAuth token has expired.
       */
      static hasExpired(token: IToken): boolean;
      /**
       * Extends Storage's default get method
       * Gets an OAuth Token after checking its expiry
       *
       * @param {string} provider Unique name of the corresponding OAuth Token.
       * @return {object} Returns the token or null if its either expired or doesn't exist.
       */
      get(provider: string): IToken;
      /**
       * Extends Storage's default add method
       * Adds a new OAuth Token after settings its expiry
       *
       * @param {string} provider Unique name of the corresponding OAuth Token.
       * @param {object} config valid Token
       * @see {@link IToken}.
       * @return {object} Returns the added token.
       */
      add(provider: string, value: IToken): IToken;
  }
  /**
   * Custom error type to handle API specific errors.
   */
  export class APIError extends CustomError {
      innerError: Error;
      /**
       * @constructor
       *
       * @param message: Error message to be propagated.
       * @param innerError: Inner error if any
      */
      constructor(message: string, innerError?: Error);
  }
  /**
   * Custom error type
   */
  export abstract class CustomError extends Error {
      name: string;
      message: string;
      innerError: Error;
      constructor(name: string, message: string, innerError?: Error);
  }
  /// <reference types="office-js" />
  /**
   * Helper exposing useful Utilities for Excel Add-ins.
   */
  export class ExcelUtilities {
      /**
       * Utility to create (or re-create) a worksheet, even if it already exists.
       * @param workbook
       * @param sheetName
       * @param clearOnly If the sheet already exists, keep it as is, and only clear its grid.
       * This results in a faster operation, and avoid a screen-update flash
       * (and the re-setting of the current selection).
       * Note: Clearing the grid does not remove floating objects like charts.
       * @returns the new worksheet
       */
      static forceCreateSheet(workbook: Excel.Workbook, sheetName: string, clearOnly?: boolean): Promise<Excel.Worksheet>;
  }
  /**
   * Custom error type to handle API specific errors.
   */
  export class DialogError extends CustomError {
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
   * An optimized size object computed based on Screen Height & Screen Width
   */
  export interface IDialogSize {
      /**
       * Width in pixels
       */
      width: number;
      /**
       * Width in percentage
       */
      width$: number;
      /**
       * Height in pixels
       */
      height: number;
      /**
       * Height in percentage
       */
      height$: number;
  }
  export class Dialog<T> {
      url: string;
      useTeamsDialog: boolean;
      /**
       * @constructor
       *
       * @param url Url to be opened in the dialog.
       * @param width Width of the dialog.
       * @param height Height of the dialog.
      */
      constructor(url?: string, width?: number, height?: number, useTeamsDialog?: boolean);
      private _result;
      readonly result: Promise<T>;
      size: IDialogSize;
      private _addinDialog();
      private _teamsDialog();
      /**
       * Close any open dialog by providing an optional message.
       * If more than one dialogs are attempted to be opened
       * an expcetion will be created.
       */
      static close(message?: any, useTeamsDialog?: boolean): void;
      private _optimizeSize(width, height);
      private _maxSize(value, max);
      private _percentage(value, max);
      private _safeParse(data);
  }
  /**
   * Helper for creating and querying Dictionaries.
   * A rudimentary alternative to ES6 Maps.
   */
  export class Dictionary<T> {
      protected items: {
          [index: string]: T;
      };
      /**
       * @constructor
       * @param {object} items Initial seed of items.
      */
      constructor(items?: {
          [index: string]: T;
      });
      /**
       * Gets an item from the dictionary.
       *
       * @param {string} key The key of the item.
       * @return {object} Returns an item if found, else returns null.
       */
      get(key: string): T;
      /**
       * Adds an item into the dictionary.
       * If the key already exists, then it will throw.
       *
       * @param {string} key The key of the item.
       * @param {object} value The item to be added.
       * @return {object} Returns the added item.
       */
      add(key: string, value: T): T;
      /**
       * Inserts an item into the dictionary.
       * If an item already exists with the same key, it will be overridden by the new value.
       *
       * @param {string} key The key of the item.
       * @param {object} value The item to be added.
       * @return {object} Returns the added item.
       */
      insert(key: string, value: T): T;
      /**
       * Removes an item from the dictionary.
       * Will throw if the key doesn't exist.
       *
       * @param {string} key The key of the item.
       * @return {object} Returns the deleted item.
       */
      remove(key: string): T;
      /**
       * Clears the dictionary.
       */
      clear(): void;
      /**
       * Check if the dictionary contains the given key.
       *
       * @param {string} key The key of the item.
       * @return {boolean} Returns true if the key was found.
       */
      contains(key: string): boolean;
      /**
       * Lists all the keys in the dictionary.
       *
       * @return {array} Returns all the keys.
       */
      keys(): string[];
      /**
       * Lists all the values in the dictionary.
       *
       * @return {array} Returns all the values.
       */
      values(): T[];
      /**
       * Get the dictionary.
       *
       * @return {object} Returns the dictionary if it contains data, null otherwise.
       */
      lookup(): {
          [key: string]: T;
      };
      /**
       * Number of items in the dictionary.
       *
       * @return {number} Returns the number of items in the dictionary.
       */
      readonly count: number;
  }
  export enum StorageType {
      LocalStorage = 0,
      SessionStorage = 1,
  }
  export interface Listener {
      subscribe(): Subscription;
      subscribe(next?: () => void, error?: (error: any) => void, complete?: () => void): Subscription;
  }
  export interface Subscription {
      /**
       * A flag to indicate whether this Subscription has already been unsubscribed.
       * @type {boolean}
       */
      closed: boolean;
      /**
       * Disposes the resources held by the subscription. May, for instance, cancel
       * an ongoing Observable execution or cancel any other type of work that
       * started when the Subscription was created.
       * @return {void}
       */
      unsubscribe(): void;
  }
  /**
   * Helper for creating and querying Local Storage or Session Storage.
   * Uses {@link Dictionary} so all the data is encapsulated in a single
   * storage namespace. Writes update the actual storage.
   */
  export class Storage<T> extends Dictionary<T> {
      container: string;
      private _type;
      private _storage;
      private readonly _current;
      /**
       * @constructor
       * @param {string} container Container name to be created in the LocalStorage.
       * @param {StorageType} type[optional] Storage Type to be used, defaults to Local Storage.
      */
      constructor(container: string, _type?: StorageType);
      /**
       * Switch the storage type.
       * Switches the storage type and then reloads the in-memory collection.
       *
       * @type {StorageType} type The desired storage to be used.
       */
      switchStorage(type: StorageType): void;
      /**
       * Add an item.
       * Extends Dictionary's implementation of add, with a save to the storage.
       */
      add(item: string, value: T): T;
      /**
       * Add or Update an item.
       * Extends Dictionary's implementation of insert, with a save to the storage.
       */
      insert(item: string, value: T): T;
      /**
       * Remove an item.
       * Extends Dictionary's implementation with a save to the storage.
       */
      remove(item: string): T;
      /**
       * Clear the storage.
       * Extends Dictionary's implementation with a save to the storage.
       */
      clear(): void;
      /**
       * Clear all storages.
       * Completely clears both the localStorage and sessionStorage.
       */
      static clearAll(): void;
      /**
       * Refreshes the storage with the current localStorage values.
       */
      load(): void;
      /**
       * Notify that the storage has changed only if the 'notify'
       * property has been subscribed to.
       */
      notify: () => Listener;
      /**
       * Synchronizes the current state to the storage.
       */
      private _sync(item, value);
  }
  /**
   * Constant strings for the host types
   */
  export const HostType: {
      WEB: string;
      ACCESS: string;
      EXCEL: string;
      ONENOTE: string;
      OUTLOOK: string;
      POWERPOINT: string;
      PROJECT: string;
      WORD: string;
  };
  /**
   * Constant strings for the host platforms
   */
  export const PlatformType: {
      IOS: string;
      MAC: string;
      OFFICE_ONLINE: string;
      PC: string;
  };
  /**
   * Helper exposing useful Utilities for Office-Add-ins.
   */
  export class Utilities {
      static readonly host: string;
      static readonly platform: string;
      /**
       * Utility to check if the code is running inside of an add-in.
       */
      static readonly isAddin: boolean;
      /**
       * Utility to print prettified errors.
       * If multiple parameters are sent then it just logs them instead.
       */
      static log(exception: Error | CustomError | string, extras?: any, ...args: any[]): void;
  }
  export class UI {
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
}