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
/**
 * Custom error type
 */
var CustomError = (function (_super) {
    __extends(CustomError, _super);
    function CustomError(name, message, innerError) {
        var _this = _super.call(this, message) || this;
        _this.name = name;
        _this.message = message;
        _this.innerError = innerError;
        if (Error.captureStackTrace) {
            Error.captureStackTrace(_this, _this.constructor);
        }
        else {
            var error = new Error();
            if (error.stack) {
                var last_part = error.stack.match(/[^\s]+$/);
                _this.stack = _this.name + " at " + last_part;
            }
        }
        return _this;
    }
    return CustomError;
}(Error));
exports.CustomError = CustomError;
//# sourceMappingURL=custom.error.js.map