/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. */
"use strict";
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
Object.defineProperty(exports, "__esModule", { value: true });
var lodash_1 = require("lodash");
var custom_error_1 = require("../errors/custom.error");
var utilities_1 = require("../helpers/utilities");
var UI = (function () {
    function UI() {
    }
    UI.notify = function () {
        var params = parseNotificationParams(arguments);
        var messageBarClasses = {
            'success': 'ms-MessageBar--success',
            'error': 'ms-MessageBar--error',
            'warning': 'ms-MessageBar--warning',
            'severe-warning': 'ms-MessageBar--severeWarning'
        };
        var messageBarTypeClass = messageBarClasses[params.type] || '';
        var paddingForPersonalityMenu = '0';
        if (utilities_1.Utilities.platform === utilities_1.PlatformType.PC) {
            paddingForPersonalityMenu = '20px';
        }
        else if (utilities_1.Utilities.platform === utilities_1.PlatformType.MAC) {
            paddingForPersonalityMenu = '40px';
        }
        var messageBannerHtml = "\n            <div class=\"office-js-helpers-notification ms-font-m ms-MessageBar " + messageBarTypeClass + "\">\n                <style>\n                    .office-js-helpers-notification {\n                        position: fixed;\n                        z-index: 2147483647;\n                        top: 0;\n                        left: 0;\n                        right: 0;\n                        width: 100%;\n                        padding: 0 0 10px 0;\n                    }\n                    .office-js-helpers-notification > div > div {\n                        padding: 10px 15px;\n                        box-sizing: border-box;\n                    }\n                    .office-js-helpers-notification pre {\n                        white-space: pre-wrap;\n                        word-wrap: break-word;\n                        margin: 0px;\n                        font-size: smaller;\n                    }\n                    .office-js-helpers-notification > button {\n                        height: 52px;\n                        width: 40px;\n                        cursor: pointer;\n                        float: right;\n                        background: transparent;\n                        border: 0;\n                        margin-left: 10px;\n                        margin-right: " + paddingForPersonalityMenu + "\n                    }\n                </style>\n                <button>\n                    <i class=\"ms-Icon ms-Icon--Clear\"></i>\n                </button>\n            </div>";
        var existingNotifications = document.getElementsByClassName('office-js-helpers-notification');
        while (existingNotifications[0]) {
            existingNotifications[0].parentNode.removeChild(existingNotifications[0]);
        }
        document.body.insertAdjacentHTML('afterbegin', messageBannerHtml);
        var notificationDiv = document.getElementsByClassName('office-js-helpers-notification')[0];
        var messageTextArea = document.createElement('div');
        notificationDiv.insertAdjacentElement('beforeend', messageTextArea);
        if (params.title) {
            var titleDiv = document.createElement('div');
            titleDiv.textContent = params.title;
            titleDiv.classList.add('ms-fontWeight-semibold');
            messageTextArea.insertAdjacentElement('beforeend', titleDiv);
        }
        params.message.split('\n').forEach(function (text) {
            var div = document.createElement('div');
            div.textContent = text;
            messageTextArea.insertAdjacentElement('beforeend', div);
        });
        if (params.moreDetails) {
            var labelDiv_1 = document.createElement('div');
            messageTextArea.insertAdjacentElement('beforeend', labelDiv_1);
            var label = document.createElement('a');
            label.setAttribute('href', 'javascript:void(0)');
            label.onclick = function () {
                document.querySelector('.office-js-helpers-notification pre')
                    .parentElement.style.display = 'block';
                labelDiv_1.style.display = 'none';
            };
            label.textContent = params.moreDetailsLabel;
            labelDiv_1.insertAdjacentElement('beforeend', label);
            var preDiv = document.createElement('div');
            preDiv.style.display = 'none';
            messageTextArea.insertAdjacentElement('beforeend', preDiv);
            var detailsDiv = document.createElement('pre');
            detailsDiv.textContent = params.moreDetails;
            preDiv.insertAdjacentElement('beforeend', detailsDiv);
        }
        document.querySelector('.office-js-helpers-notification > button')
            .onclick = function () {
            notificationDiv.parentNode.removeChild(notificationDiv);
        };
    };
    return UI;
}());
exports.UI = UI;
function parseNotificationParams(params) {
    try {
        var defaults = {
            title: null,
            type: 'default',
            moreDetails: null,
            moreDetailsLabel: 'Additional details...'
        };
        switch (params.length) {
            case 1: {
                if (lodash_1.isError(params[0])) {
                    return __assign({}, defaults, { title: 'Error', type: 'error' }, getErrorDetails(params[0]));
                }
                if (lodash_1.isString(params[0])) {
                    return __assign({}, defaults, { message: params[0] });
                }
                if (lodash_1.isObject(params[0])) {
                    var customParams = params[0];
                    if (!lodash_1.isString(customParams.message)) {
                        throw new Error();
                    }
                    return __assign({}, defaults, { title: customParams.title || defaults.title, message: customParams.message, type: customParams.type || defaults.type });
                }
                throw new Error();
            }
            case 2: {
                if (lodash_1.isString(params[0])) {
                    if (lodash_1.isError(params[1])) {
                        return __assign({}, defaults, { title: params[0] }, getErrorDetails(params[1]));
                    }
                    if (lodash_1.isString(params[1])) {
                        return __assign({}, defaults, { title: params[0], message: params[1] });
                    }
                }
                else if (lodash_1.isError(params[0]) && lodash_1.isObject(params[1])) {
                    var customParams = params[1];
                    var result = __assign({}, defaults, getErrorDetails(params[0]), { moreDetailsLabel: customParams.moreDetailsLabel || defaults.moreDetailsLabel });
                    result.title = customParams.title || result.title;
                    result.message = customParams.message || result.message;
                    return result;
                }
                throw new Error();
            }
            case 3: {
                if (!(lodash_1.isString(params[0]) && lodash_1.isString(params[2]))) {
                    throw new Error();
                }
                if (!lodash_1.isString(params[1])) {
                    throw new Error();
                }
                return __assign({}, defaults, { title: params[0], message: params[1], type: params[2] });
            }
            default:
                throw new Error();
        }
    }
    catch (e) {
        throw new Error('Invalid parameters passed to "notify" function');
    }
}
function getErrorDetails(error) {
    var moreDetails;
    var innerException = error;
    if (error instanceof custom_error_1.CustomError) {
        innerException = error.innerError;
    }
    if (window.OfficeExtension && innerException instanceof OfficeExtension.Error) {
        moreDetails = JSON.stringify(error.debugInfo, null, 4);
    }
    return {
        type: 'error',
        message: error.toString(),
        moreDetails: moreDetails
    };
}
//# sourceMappingURL=ui.js.map