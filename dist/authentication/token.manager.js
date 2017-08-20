/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. */
"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var storage_1 = require("../helpers/storage");
/**
 * Helper for caching and managing OAuth Tokens.
 */
var TokenStorage = (function (_super) {
    __extends(TokenStorage, _super);
    /**
     * @constructor
    */
    function TokenStorage() {
        return _super.call(this, 'OAuth2Tokens') || this;
    }
    /**
     * Compute the expiration date based on the expires_in field in a OAuth token.
     */
    TokenStorage.setExpiry = function (token) {
        var expire = function (seconds) { return seconds == null ? null : new Date(new Date().getTime() + ~~seconds * 1000); };
        if (!(token == null) && token.expires_at == null) {
            token.expires_at = expire(token.expires_in);
        }
    };
    /**
     * Check if an OAuth token has expired.
     */
    TokenStorage.hasExpired = function (token) {
        if (token == null) {
            return true;
        }
        if (token.expires_at == null) {
            return false;
        }
        else {
            // If the token was stored, it's Date type property was stringified, so it needs to be converted back to Date.
            token.expires_at = token.expires_at instanceof Date ? token.expires_at : new Date(token.expires_at);
            return token.expires_at.getTime() - new Date().getTime() < 0;
        }
    };
    /**
     * Extends Storage's default get method
     * Gets an OAuth Token after checking its expiry
     *
     * @param {string} provider Unique name of the corresponding OAuth Token.
     * @return {object} Returns the token or null if its either expired or doesn't exist.
     */
    TokenStorage.prototype.get = function (provider) {
        var token = _super.prototype.get.call(this, provider);
        if (token == null) {
            return token;
        }
        var expired = TokenStorage.hasExpired(token);
        if (expired) {
            _super.prototype.remove.call(this, provider);
            return null;
        }
        else {
            return token;
        }
    };
    /**
     * Extends Storage's default add method
     * Adds a new OAuth Token after settings its expiry
     *
     * @param {string} provider Unique name of the corresponding OAuth Token.
     * @param {object} config valid Token
     * @see {@link IToken}.
     * @return {object} Returns the added token.
     */
    TokenStorage.prototype.add = function (provider, value) {
        value.provider = provider;
        TokenStorage.setExpiry(value);
        return _super.prototype.insert.call(this, provider, value);
    };
    return TokenStorage;
}(storage_1.Storage));
exports.TokenStorage = TokenStorage;
//# sourceMappingURL=token.manager.js.map