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
var utilities_1 = require("./utilities");
var custom_error_1 = require("../errors/custom.error");
/**
 * Custom error type to handle API specific errors.
 */
var DialogError = (function (_super) {
    __extends(DialogError, _super);
    /**
     * @constructor
     *
     * @param message Error message to be propagated.
     * @param state OAuth state if available.
    */
    function DialogError(message, innerError) {
        var _this = _super.call(this, 'DialogError', message, innerError) || this;
        _this.innerError = innerError;
        return _this;
    }
    return DialogError;
}(custom_error_1.CustomError));
exports.DialogError = DialogError;
var Dialog = (function () {
    /**
     * @constructor
     *
     * @param url Url to be opened in the dialog.
     * @param width Width of the dialog.
     * @param height Height of the dialog.
    */
    function Dialog(url, width, height, useTeamsDialog) {
        if (url === void 0) { url = location.origin; }
        if (width === void 0) { width = 1024; }
        if (height === void 0) { height = 768; }
        if (useTeamsDialog === void 0) { useTeamsDialog = false; }
        this.url = url;
        this.useTeamsDialog = useTeamsDialog;
        if (!(/^https/.test(url))) {
            throw new DialogError('URL has to be loaded over HTTPS.');
        }
        this.size = this._optimizeSize(width, height);
    }
    Object.defineProperty(Dialog.prototype, "result", {
        get: function () {
            if (this._result == null) {
                this._result = this.useTeamsDialog ? this._teamsDialog() : this._addinDialog();
            }
            return this._result;
        },
        enumerable: true,
        configurable: true
    });
    Dialog.prototype._addinDialog = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            Office.context.ui.displayDialogAsync(_this.url, { width: _this.size.width$, height: _this.size.height$ }, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    throw new DialogError(result.error.message);
                }
                else {
                    var dialog_1 = result.value;
                    dialog_1.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
                        try {
                            var result_1 = _this._safeParse(args.message);
                            if (result_1.parse) {
                                resolve(_this._safeParse(result_1.value));
                            }
                            else {
                                resolve(result_1.value);
                            }
                        }
                        catch (exception) {
                            reject(new DialogError('An unexpected error in the dialog has occured.', exception));
                        }
                        finally {
                            dialog_1.close();
                        }
                    });
                    dialog_1.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
                        try {
                            reject(new DialogError(args.message, args.error));
                        }
                        catch (exception) {
                            reject(new DialogError('An unexpected error in the dialog has occured.', exception));
                        }
                        finally {
                            dialog_1.close();
                        }
                    });
                }
            });
        });
    };
    Dialog.prototype._teamsDialog = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            try {
                microsoftTeams.initialize();
            }
            catch (e) {
            }
            microsoftTeams.authentication.authenticate({
                url: _this.url,
                width: _this.size.width,
                height: _this.size.height,
                failureCallback: function (exception) { return reject(new DialogError('Error while launching dialog', exception)); },
                successCallback: function (message) { return resolve(message); }
            });
        });
    };
    /**
     * Close any open dialog by providing an optional message.
     * If more than one dialogs are attempted to be opened
     * an expcetion will be created.
     */
    Dialog.close = function (message, useTeamsDialog) {
        if (useTeamsDialog === void 0) { useTeamsDialog = false; }
        var parse = false;
        var value = message;
        if ((!(value == null)) && typeof value === 'object') {
            parse = true;
            value = JSON.stringify(value);
        }
        else if (typeof message === 'function') {
            throw new DialogError('Invalid message. Cannot pass functions as arguments');
        }
        try {
            if (useTeamsDialog) {
                try {
                    microsoftTeams.initialize();
                }
                catch (e) {
                }
                microsoftTeams.authentication.notifySuccess(JSON.stringify({ parse: parse, value: value }));
            }
            else if (utilities_1.Utilities.isAddin) {
                Office.context.ui.messageParent(JSON.stringify({ parse: parse, value: value }));
            }
        }
        catch (error) {
            throw new DialogError('Cannot close dialog', error);
        }
    };
    Dialog.prototype._optimizeSize = function (width, height) {
        var screenWidth = window.screen.width;
        var screenHeight = window.screen.height;
        var optimizedWidth = this._maxSize(width, screenWidth);
        var optimizedHeight = this._maxSize(height, screenHeight);
        return {
            width$: this._percentage(optimizedWidth, screenWidth),
            height$: this._percentage(optimizedHeight, screenHeight),
            width: optimizedWidth,
            height: optimizedHeight
        };
    };
    Dialog.prototype._maxSize = function (value, max) {
        return value < (max - 30) ? value : max - 30;
    };
    ;
    Dialog.prototype._percentage = function (value, max) {
        return (value * 100 / max);
    };
    Dialog.prototype._safeParse = function (data) {
        try {
            var result = JSON.parse(data);
            return result;
        }
        catch (e) {
            return data;
        }
    };
    return Dialog;
}());
exports.Dialog = Dialog;
//# sourceMappingURL=dialog.js.map