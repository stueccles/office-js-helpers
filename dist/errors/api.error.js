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
var custom_error_1 = require("./custom.error");
/**
 * Custom error type to handle API specific errors.
 */
var APIError = (function (_super) {
    __extends(APIError, _super);
    /**
     * @constructor
     *
     * @param message: Error message to be propagated.
     * @param innerError: Inner error if any
    */
    function APIError(message, innerError) {
        var _this = _super.call(this, 'APIError', message, innerError) || this;
        _this.innerError = innerError;
        return _this;
    }
    return APIError;
}(custom_error_1.CustomError));
exports.APIError = APIError;
//# sourceMappingURL=api.error.js.map