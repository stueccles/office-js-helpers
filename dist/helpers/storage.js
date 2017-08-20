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
var debounce = require("lodash/debounce");
var dictionary_1 = require("./dictionary");
var md5 = require("crypto-js/md5");
var Observable_1 = require("rxjs/Observable");
var StorageType;
(function (StorageType) {
    StorageType[StorageType["LocalStorage"] = 0] = "LocalStorage";
    StorageType[StorageType["SessionStorage"] = 1] = "SessionStorage";
})(StorageType = exports.StorageType || (exports.StorageType = {}));
/**
 * Helper for creating and querying Local Storage or Session Storage.
 * Uses {@link Dictionary} so all the data is encapsulated in a single
 * storage namespace. Writes update the actual storage.
 */
var Storage = (function (_super) {
    __extends(Storage, _super);
    /**
     * @constructor
     * @param {string} container Container name to be created in the LocalStorage.
     * @param {StorageType} type[optional] Storage Type to be used, defaults to Local Storage.
    */
    function Storage(container, _type) {
        var _this = _super.call(this) || this;
        _this.container = container;
        _this._type = _type;
        _this._storage = null;
        /**
         * Notify that the storage has changed only if the 'notify'
         * property has been subscribed to.
         */
        _this.notify = function () {
            return new Observable_1.Observable(function (observer) {
                /* Determine the initial count and hash for this loop */
                var lastCount = _this.count;
                var lastHash = md5(JSON.stringify(_this.items)).toString();
                /* Begin the polling at 300ms */
                var pollInterval = setInterval(function () {
                    try {
                        _this.load();
                        /* If the last count isn't the same as the current count */
                        if (_this.count !== lastCount) {
                            lastCount = _this.count;
                            observer.next();
                        }
                        else {
                            var hash = md5(JSON.stringify(_this.items)).toString();
                            /* If the last hash isn't the same as the current hash */
                            if (hash !== lastHash) {
                                lastHash = hash;
                                observer.next();
                            }
                        }
                    }
                    catch (e) {
                        observer.error(e);
                    }
                }, 300);
                /* Debounced listener to localStorage events given that they fire any change */
                var debouncedUpdate = debounce(function (event) {
                    try {
                        clearInterval(pollInterval);
                        /* If the change is on the current container */
                        if (event.key === _this.container) {
                            _this.load();
                            observer.next();
                        }
                    }
                    catch (e) {
                        observer.error(e);
                    }
                }, 300);
                window.addEventListener('storage', debouncedUpdate, false);
                /* Teardown */
                return function () {
                    if (pollInterval) {
                        clearInterval(pollInterval);
                    }
                    window.removeEventListener('storage', debouncedUpdate, false);
                };
            });
        };
        _this._type = _this._type || StorageType.LocalStorage;
        _this.switchStorage(_this._type);
        return _this;
    }
    Object.defineProperty(Storage.prototype, "_current", {
        get: function () {
            return JSON.parse(this._storage.getItem(this.container));
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Switch the storage type.
     * Switches the storage type and then reloads the in-memory collection.
     *
     * @type {StorageType} type The desired storage to be used.
     */
    Storage.prototype.switchStorage = function (type) {
        this._storage = type === StorageType.LocalStorage ? localStorage : sessionStorage;
        if (!this._storage.hasOwnProperty(this.container)) {
            this._storage[this.container] = null;
        }
        this.load();
    };
    /**
     * Add an item.
     * Extends Dictionary's implementation of add, with a save to the storage.
     */
    Storage.prototype.add = function (item, value) {
        _super.prototype.add.call(this, item, value);
        this._sync(item, value);
        return value;
    };
    /**
     * Add or Update an item.
     * Extends Dictionary's implementation of insert, with a save to the storage.
     */
    Storage.prototype.insert = function (item, value) {
        _super.prototype.insert.call(this, item, value);
        this._sync(item, value);
        return value;
    };
    /**
     * Remove an item.
     * Extends Dictionary's implementation with a save to the storage.
     */
    Storage.prototype.remove = function (item) {
        var value = _super.prototype.remove.call(this, item);
        this._sync(item, null);
        return value;
    };
    /**
     * Clear the storage.
     * Extends Dictionary's implementation with a save to the storage.
     */
    Storage.prototype.clear = function () {
        _super.prototype.clear.call(this);
        this._storage.removeItem(this.container);
    };
    /**
     * Clear all storages.
     * Completely clears both the localStorage and sessionStorage.
     */
    Storage.clearAll = function () {
        window.localStorage.clear();
        window.sessionStorage.clear();
    };
    /**
     * Refreshes the storage with the current localStorage values.
     */
    Storage.prototype.load = function () {
        var items = __assign({}, this.items, this._current);
        this.items = items;
    };
    /**
     * Synchronizes the current state to the storage.
     */
    Storage.prototype._sync = function (item, value) {
        var items = __assign({}, this._current, this.items);
        if (value == null) {
            delete items[item];
        }
        this._storage.setItem(this.container, JSON.stringify(items));
        this.items = items;
    };
    return Storage;
}(dictionary_1.Dictionary));
exports.Storage = Storage;
//# sourceMappingURL=storage.js.map