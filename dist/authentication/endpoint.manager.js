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
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
Object.defineProperty(exports, "__esModule", { value: true });
var storage_1 = require("../helpers/storage");
exports.DefaultEndpoints = {
    Google: 'Google',
    Microsoft: 'Microsoft',
    Facebook: 'Facebook',
    AzureAD: 'AzureAD'
};
/**
 * Helper for creating and registering OAuth Endpoints.
 */
var EndpointStorage = (function (_super) {
    __extends(EndpointStorage, _super);
    /**
     * @constructor
    */
    function EndpointStorage() {
        return _super.call(this, 'OAuth2Endpoints') || this;
    }
    /**
     * Extends Storage's default add method.
     * Registers a new OAuth Endpoint.
     *
     * @param {string} provider Unique name for the registered OAuth Endpoint.
     * @param {object} config Valid Endpoint configuration.
     * @see {@link IEndpointConfiguration}.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.prototype.add = function (provider, config) {
        if (config.redirectUrl == null) {
            config.redirectUrl = window.location.origin;
        }
        config.provider = provider;
        return _super.prototype.insert.call(this, provider, config);
    };
    /**
     * Register Google Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Google App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.prototype.registerGoogleAuth = function (clientId, overrides) {
        var defaults = {
            clientId: clientId,
            baseUrl: 'https://accounts.google.com',
            authorizeUrl: '/o/oauth2/v2/auth',
            resource: 'https://www.googleapis.com',
            responseType: 'token',
            scope: 'https://www.googleapis.com/auth/plus.me',
            state: true
        };
        var config = __assign({}, defaults, overrides);
        return this.add(exports.DefaultEndpoints.Google, config);
    };
    ;
    /**
     * Register Microsoft Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Microsoft App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.prototype.registerMicrosoftAuth = function (clientId, overrides) {
        var defaults = {
            clientId: clientId,
            baseUrl: 'https://login.microsoftonline.com/common/oauth2/v2.0',
            authorizeUrl: '/authorize',
            responseType: 'token',
            scope: 'https://graph.microsoft.com/user.read',
            extraQueryParameters: {
                response_mode: 'fragment'
            },
            nonce: true,
            state: true
        };
        var config = __assign({}, defaults, overrides);
        this.add(exports.DefaultEndpoints.Microsoft, config);
    };
    ;
    /**
     * Register Facebook Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the Facebook App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.prototype.registerFacebookAuth = function (clientId, overrides) {
        var defaults = {
            clientId: clientId,
            baseUrl: 'https://www.facebook.com',
            authorizeUrl: '/dialog/oauth',
            resource: 'https://graph.facebook.com',
            responseType: 'token',
            scope: 'public_profile',
            nonce: true,
            state: true
        };
        var config = __assign({}, defaults, overrides);
        this.add(exports.DefaultEndpoints.Facebook, config);
    };
    ;
    /**
     * Register AzureAD Implicit OAuth.
     * If overrides is left empty, the default scope is limited to basic profile information.
     *
     * @param {string} clientId ClientID for the AzureAD App.
     * @param {string} tenant Tenant for the AzureAD App.
     * @param {object} config Valid Endpoint configuration to override the defaults.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.prototype.registerAzureADAuth = function (clientId, tenant, overrides) {
        var defaults = {
            clientId: clientId,
            baseUrl: "https://login.windows.net/" + tenant,
            authorizeUrl: '/oauth2/authorize',
            resource: 'https://graph.microsoft.com',
            responseType: 'token',
            nonce: true,
            state: true
        };
        var config = __assign({}, defaults, overrides);
        this.add(exports.DefaultEndpoints.AzureAD, config);
    };
    ;
    /**
     * Helper to generate the OAuth login url.
     *
     * @param {object} config Valid Endpoint configuration.
     * @return {object} Returns the added endpoint.
     */
    EndpointStorage.getLoginParams = function (endpointConfig) {
        var scope = (endpointConfig.scope) ? encodeURIComponent(endpointConfig.scope) : null;
        var resource = (endpointConfig.resource) ? encodeURIComponent(endpointConfig.resource) : null;
        var state = endpointConfig.state && EndpointStorage.generateCryptoSafeRandom();
        var nonce = endpointConfig.nonce && EndpointStorage.generateCryptoSafeRandom();
        var urlSegments = [
            "response_type=" + endpointConfig.responseType,
            "client_id=" + encodeURIComponent(endpointConfig.clientId),
            "redirect_uri=" + encodeURIComponent(endpointConfig.redirectUrl)
        ];
        if (scope) {
            urlSegments.push("scope=" + scope);
        }
        if (resource) {
            urlSegments.push("resource=" + resource);
        }
        if (state) {
            urlSegments.push("state=" + state);
        }
        if (nonce) {
            urlSegments.push("nonce=" + nonce);
        }
        if (endpointConfig.extraQueryParameters) {
            for (var _i = 0, _a = Object.keys(endpointConfig.extraQueryParameters); _i < _a.length; _i++) {
                var param = _a[_i];
                urlSegments.push(param + "=" + encodeURIComponent(endpointConfig.extraQueryParameters[param]));
            }
        }
        return {
            url: "" + endpointConfig.baseUrl + endpointConfig.authorizeUrl + "?" + urlSegments.join('&'),
            state: state
        };
    };
    EndpointStorage.generateCryptoSafeRandom = function () {
        var random = new Uint32Array(1);
        if ('msCrypto' in window) {
            window.msCrypto.getRandomValues(random);
        }
        else if ('crypto' in window) {
            window.crypto.getRandomValues(random);
        }
        else {
            throw new Error('The platform doesn\'t support generation of cryptographically safe randoms. Please disable the state flag and try again.');
        }
        return random[0];
    };
    return EndpointStorage;
}(storage_1.Storage));
exports.EndpointStorage = EndpointStorage;
//# sourceMappingURL=endpoint.manager.js.map