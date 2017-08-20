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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t;
    return { next: verb(0), "throw": verb(1), "return": verb(2) };
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var endpoint_manager_1 = require("./endpoint.manager");
var token_manager_1 = require("./token.manager");
var utilities_1 = require("../helpers/utilities");
var dialog_1 = require("../helpers/dialog");
var custom_error_1 = require("../errors/custom.error");
/**
 * Custom error type to handle OAuth specific errors.
 */
var AuthError = (function (_super) {
    __extends(AuthError, _super);
    /**
     * @constructor
     *
     * @param message Error message to be propagated.
     * @param state OAuth state if available.
    */
    function AuthError(message, innerError) {
        var _this = _super.call(this, 'AuthError', message, innerError) || this;
        _this.innerError = innerError;
        return _this;
    }
    return AuthError;
}(custom_error_1.CustomError));
exports.AuthError = AuthError;
/**
 * Helper for performing Implicit OAuth Authentication with registered endpoints.
 */
var Authenticator = (function () {
    /**
     * @constructor
     *
     * @param endpoints Depends on an instance of EndpointStorage.
     * @param tokens Depends on an instance of TokenStorage.
    */
    function Authenticator(endpoints, tokens) {
        this.endpoints = endpoints;
        this.tokens = tokens;
        if (endpoints == null) {
            this.endpoints = new endpoint_manager_1.EndpointStorage();
        }
        if (tokens == null) {
            this.tokens = new token_manager_1.TokenStorage();
        }
    }
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
    Authenticator.prototype.authenticate = function (provider, force, useMicrosoftTeams) {
        if (force === void 0) { force = false; }
        if (useMicrosoftTeams === void 0) { useMicrosoftTeams = false; }
        var token = this.tokens.get(provider);
        var hasTokenExpired = token_manager_1.TokenStorage.hasExpired(token);
        if (!hasTokenExpired && !force) {
            return Promise.resolve(token);
        }
        if (useMicrosoftTeams) {
            return this._openAuthDialog(provider, true);
        }
        else if (utilities_1.Utilities.isAddin) {
            return this._openAuthDialog(provider, false);
        }
        else {
            return this._openInWindowPopup(provider);
        }
    };
    /**
     * Check if the currrent url is running inside of a Dialog that contains an access_token or code or error.
     * If true then it calls messageParent by extracting the token information, thereby closing the dialog.
     * Otherwise, the caller should proceed with normal initialization of their application.
     *
     * @return {boolean}
     * Returns false if the code is running inside of a dialog without the required information
     * or is not running inside of a dialog at all.
     */
    Authenticator.isAuthDialog = function (useMicrosoftTeams) {
        if (useMicrosoftTeams === void 0) { useMicrosoftTeams = false; }
        if (useMicrosoftTeams === false && !utilities_1.Utilities.isAddin) {
            return false;
        }
        else {
            if (!/(access_token|code|error)/gi.test(location.href)) {
                return false;
            }
            dialog_1.Dialog.close(location.href, useMicrosoftTeams);
            return true;
        }
    };
    /**
     * Extract the token from the URL
     *
     * @param {string} url The url to extract the token from.
     * @param {string} exclude Exclude a particlaur string from the url, such as a query param or specific substring.
     * @param {string} delimiter[optional] Delimiter used by OAuth provider to mark the beginning of token response. Defaults to #.
     * @return {object} Returns the extracted token.
     */
    Authenticator.getUrlParams = function (url, exclude, delimiter) {
        if (url === void 0) { url = location.href; }
        if (exclude === void 0) { exclude = location.origin; }
        if (delimiter === void 0) { delimiter = '#'; }
        if (exclude) {
            url = url.replace(exclude, '');
        }
        var _a = url.split(delimiter), left = _a[0], right = _a[1];
        var tokenString = right == null ? left : right;
        if (tokenString.indexOf('?') !== -1) {
            tokenString = tokenString.split('?')[1];
        }
        return Authenticator.extractParams(tokenString);
    };
    Authenticator.extractParams = function (segment) {
        if (segment == null || segment.trim() === '') {
            return null;
        }
        var params = {}, regex = /([^&=]+)=([^&]*)/g, matchParts;
        while ((matchParts = regex.exec(segment)) !== null) {
            /* Fixes bugs when the state parameters contains a / before them */
            if (matchParts[1] === '/state') {
                matchParts[1] = matchParts[1].replace('/', '');
            }
            params[decodeURIComponent(matchParts[1])] = decodeURIComponent(matchParts[2]);
        }
        return params;
    };
    Authenticator.prototype._openAuthDialog = function (provider, useMicrosoftTeams) {
        return __awaiter(this, void 0, void 0, function () {
            var endpoint, _a, state, url, redirectUrl;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        endpoint = this.endpoints.get(provider);
                        if (endpoint == null) {
                            return [2 /*return*/, Promise.reject(new AuthError("No such registered endpoint: " + provider + " could be found."))];
                        }
                        _a = endpoint_manager_1.EndpointStorage.getLoginParams(endpoint), state = _a.state, url = _a.url;
                        return [4 /*yield*/, new dialog_1.Dialog(url, 1024, 768, useMicrosoftTeams).result];
                    case 1:
                        redirectUrl = _b.sent();
                        /** Try and extract the result and pass it along */
                        return [2 /*return*/, this._handleTokenResult(redirectUrl, endpoint, state)];
                }
            });
        });
    };
    Authenticator.prototype._openInWindowPopup = function (provider) {
        var _this = this;
        /** Get the endpoint configuration for the given provider and verify that it exists. */
        var endpoint = this.endpoints.get(provider);
        if (endpoint == null) {
            return Promise.reject(new AuthError("No such registered endpoint: " + provider + " could be found."));
        }
        var _a = endpoint_manager_1.EndpointStorage.getLoginParams(endpoint), state = _a.state, url = _a.url;
        var windowFeatures = "width=" + 1024 + ",height=" + 768 + ",menubar=no,toolbar=no,location=no,resizable=yes,scrollbars=yes,status=no";
        var popupWindow = window.open(url, endpoint.provider.toUpperCase(), windowFeatures);
        return new Promise(function (resolve, reject) {
            try {
                var POLL_INTERVAL = 400;
                var interval_1 = setInterval(function () {
                    try {
                        if (popupWindow.document.URL.indexOf(endpoint.redirectUrl) !== -1) {
                            clearInterval(interval_1);
                            popupWindow.close();
                            return resolve(_this._handleTokenResult(popupWindow.document.URL, endpoint, state));
                        }
                    }
                    catch (exception) {
                        if (!popupWindow) {
                            clearInterval(interval_1);
                            return reject(new AuthError('Popup window was closed'));
                        }
                    }
                }, POLL_INTERVAL);
            }
            catch (exception) {
                popupWindow.close();
                return reject(new AuthError('Unexpected error occured while creating popup'));
            }
        });
    };
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
    Authenticator.prototype._exchangeCodeForToken = function (endpoint, data, headers) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (endpoint.tokenUrl == null) {
                console.warn("We couldn't exchange the received code for an access_token.\n                    The value returned is not an access_token.\n                    Please set the tokenUrl property or refer to our docs.");
                return resolve(data);
            }
            var xhr = new XMLHttpRequest();
            xhr.open('POST', endpoint.tokenUrl);
            xhr.setRequestHeader('Accept', 'application/json');
            xhr.setRequestHeader('Content-Type', 'application/json');
            if (endpoint.grantType != null) {
                data.grantType = endpoint.grantType;
            }
            for (var header in headers) {
                if (header === 'Accept' || header === 'Content-Type') {
                    continue;
                }
                xhr.setRequestHeader(header, headers[header]);
            }
            var extraHeaders = endpoint.extraHeaders;
            if (extraHeaders != null) {
                for (var header in extraHeaders) {
                    xhr.setRequestHeader(header, extraHeaders[header]);
                }
            }
            xhr.onerror = function () {
                return reject(new AuthError('Unable to send request due to a Network error'));
            };
            xhr.onload = function () {
                try {
                    if (xhr.status === 200) {
                        var json = JSON.parse(xhr.responseText);
                        if (json == null) {
                            return reject(new AuthError('No access_token or code could be parsed.'));
                        }
                        else if ('access_token' in json) {
                            _this.tokens.add(endpoint.provider, json);
                            return resolve(json);
                        }
                        else {
                            return reject(new AuthError(json.error, json.state));
                        }
                    }
                    else if (xhr.status !== 200) {
                        return reject(new AuthError('Request failed. ' + xhr.response));
                    }
                }
                catch (e) {
                    return reject(new AuthError('An error occured while parsing the response'));
                }
            };
            xhr.send(JSON.stringify(data));
        });
    };
    Authenticator.prototype._handleTokenResult = function (redirectUrl, endpoint, state) {
        var result = Authenticator.getUrlParams(redirectUrl, endpoint.redirectUrl);
        if (result == null) {
            throw new AuthError('No access_token or code could be parsed.');
        }
        else if (endpoint.state && +result.state !== state) {
            throw new AuthError('State couldn\'t be verified');
        }
        else if ('code' in result) {
            return this._exchangeCodeForToken(endpoint, result);
        }
        else if ('access_token' in result) {
            return this.tokens.add(endpoint.provider, result);
        }
        else {
            throw new AuthError(result.error);
        }
    };
    return Authenticator;
}());
exports.Authenticator = Authenticator;
//# sourceMappingURL=authenticator.js.map