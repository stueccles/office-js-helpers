/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. */
"use strict";
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
var api_error_1 = require("../errors/api.error");
/**
 * Helper exposing useful Utilities for Excel Add-ins.
 */
var ExcelUtilities = (function () {
    function ExcelUtilities() {
    }
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
    ExcelUtilities.forceCreateSheet = function (workbook, sheetName, clearOnly) {
        return __awaiter(this, void 0, void 0, function () {
            var sheet;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (workbook == null && typeof workbook !== typeof Excel.Workbook) {
                            throw new api_error_1.APIError('Invalid workbook parameter.');
                        }
                        if (sheetName == null || sheetName.trim() === '') {
                            throw new api_error_1.APIError('Sheet name cannot be blank.');
                        }
                        if (sheetName.length > 31) {
                            throw new api_error_1.APIError('Sheet name cannot be greater than 31 characters.');
                        }
                        if (!clearOnly) return [3 /*break*/, 2];
                        return [4 /*yield*/, createOrClear(workbook.context, workbook, sheetName)];
                    case 1:
                        sheet = _a.sent();
                        return [3 /*break*/, 4];
                    case 2: return [4 /*yield*/, recreateFromScratch(workbook.context, workbook, sheetName)];
                    case 3:
                        sheet = _a.sent();
                        _a.label = 4;
                    case 4: 
                    // To work around an issue with Office Online (tracked by the API team), it is
                    // currently necessary to do a `context.sync()` before any call to `sheet.activate()`.
                    // So to be safe, in case the caller of this helper method decides to immediately
                    // turn around and call `sheet.activate()`, call `sync` before returning the sheet.
                    return [4 /*yield*/, workbook.context.sync()];
                    case 5:
                        // To work around an issue with Office Online (tracked by the API team), it is
                        // currently necessary to do a `context.sync()` before any call to `sheet.activate()`.
                        // So to be safe, in case the caller of this helper method decides to immediately
                        // turn around and call `sheet.activate()`, call `sync` before returning the sheet.
                        _a.sent();
                        return [2 /*return*/, sheet];
                }
            });
        });
    };
    return ExcelUtilities;
}());
exports.ExcelUtilities = ExcelUtilities;
/**
 * Helpers
 */
function createOrClear(context, workbook, sheetName) {
    return __awaiter(this, void 0, void 0, function () {
        var existingSheet, oldSheet, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.4)) return [3 /*break*/, 2];
                    existingSheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                    return [4 /*yield*/, context.sync()];
                case 1:
                    _a.sent();
                    if (existingSheet.isNullObject) {
                        return [2 /*return*/, context.workbook.worksheets.add(sheetName)];
                    }
                    else {
                        existingSheet.getRange().clear();
                        return [2 /*return*/, existingSheet];
                    }
                    return [3 /*break*/, 7];
                case 2: 
                // Flush anything already in the queue, so as to scope the error handling logic below.
                return [4 /*yield*/, context.sync()];
                case 3:
                    // Flush anything already in the queue, so as to scope the error handling logic below.
                    _a.sent();
                    _a.label = 4;
                case 4:
                    _a.trys.push([4, 6, , 7]);
                    oldSheet = workbook.worksheets.getItem(sheetName);
                    oldSheet.getRange().clear();
                    return [4 /*yield*/, context.sync()];
                case 5:
                    _a.sent();
                    return [2 /*return*/, oldSheet];
                case 6:
                    error_1 = _a.sent();
                    if (error_1 instanceof OfficeExtension.Error && error_1.code === Excel.ErrorCodes.itemNotFound) {
                        // This is an expected case where the sheet didn't exist. Create it now.
                        return [2 /*return*/, workbook.worksheets.add(sheetName)];
                    }
                    else {
                        throw new api_error_1.APIError('Unexpected error while trying to delete sheet.', error_1);
                    }
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    });
}
function recreateFromScratch(context, workbook, sheetName) {
    return __awaiter(this, void 0, void 0, function () {
        var newSheet, oldSheet, error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    newSheet = workbook.worksheets.add();
                    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.4)) return [3 /*break*/, 1];
                    context.workbook.worksheets.getItemOrNullObject(sheetName).delete();
                    return [3 /*break*/, 6];
                case 1: 
                // Flush anything already in the queue, so as to scope the error handling logic below.
                return [4 /*yield*/, context.sync()];
                case 2:
                    // Flush anything already in the queue, so as to scope the error handling logic below.
                    _a.sent();
                    _a.label = 3;
                case 3:
                    _a.trys.push([3, 5, , 6]);
                    oldSheet = workbook.worksheets.getItem(sheetName);
                    oldSheet.delete();
                    return [4 /*yield*/, context.sync()];
                case 4:
                    _a.sent();
                    return [3 /*break*/, 6];
                case 5:
                    error_2 = _a.sent();
                    if (error_2 instanceof OfficeExtension.Error && error_2.code === Excel.ErrorCodes.itemNotFound) {
                        // This is an expected case where the sheet didn't exist. Hence no-op.
                    }
                    else {
                        throw new api_error_1.APIError('Unexpected error while trying to delete sheet.', error_2);
                    }
                    return [3 /*break*/, 6];
                case 6:
                    newSheet.name = sheetName;
                    return [2 /*return*/, newSheet];
            }
        });
    });
}
//# sourceMappingURL=utilities.js.map