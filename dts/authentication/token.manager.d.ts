import { Storage } from '../helpers/storage';
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
export declare class TokenStorage extends Storage<IToken> {
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
