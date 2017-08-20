"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. */
var custom_error_1 = require("../errors/custom.error");
/**
 * Constant strings for the host types
 */
exports.HostType = {
    WEB: 'WEB',
    ACCESS: 'ACCESS',
    EXCEL: 'EXCEL',
    ONENOTE: 'ONENOTE',
    OUTLOOK: 'OUTLOOK',
    POWERPOINT: 'POWERPOINT',
    PROJECT: 'PROJECT',
    WORD: 'WORD'
};
/**
 * Constant strings for the host platforms
 */
exports.PlatformType = {
    IOS: 'IOS',
    MAC: 'MAC',
    OFFICE_ONLINE: 'OFFICE_ONLINE',
    PC: 'PC'
};
/*
* Retrieves host info using a workaround that utilizes the internals of the
* Office.js library. Such workarounds should be avoided, as they can lead to
* a break in behavior, if the internals are ever changed. In this case, however,
* Office.js will soon be delivering a new API to provide the host and platform
* information.
*/
function getHostInfo() {
    // A forthcoming API (partially rolled-out) will expose the host and platform info natively
    // when queried from within an add-in.
    // If the platform already exposes that info, then just return it
    // (but only after massaging it to fit the return types expected by this function)
    var hasContext = window['Office'] && window['Office'].context;
    var context = hasContext ? window['Office'].context : {};
    if (context.host && context.platform) {
        return {
            host: convertHostValue(context.host),
            platform: convertPlatformValue(context.platform)
        };
    }
    return useHostInfoFallbackLogic();
}
;
function useHostInfoFallbackLogic() {
    try {
        if (window.sessionStorage == null) {
            throw new Error("Session Storage isn't supported");
        }
        var hostInfoValue = window.sessionStorage['hostInfoValue'];
        var _a = hostInfoValue.split('$'), hostRaw = _a[0], platformRaw = _a[1], extras = _a[2];
        // Older hosts used "|", so check for that as well:
        if (extras == null) {
            _b = hostInfoValue.split('|'), hostRaw = _b[0], platformRaw = _b[1];
        }
        var host = hostRaw.toUpperCase() || 'WEB';
        var platform = null;
        if (Utilities.host !== exports.HostType.WEB) {
            var platforms = {
                'IOS': exports.PlatformType.IOS,
                'MAC': exports.PlatformType.MAC,
                'WEB': exports.PlatformType.OFFICE_ONLINE,
                'WIN32': exports.PlatformType.PC
            };
            platform = platforms[platformRaw.toUpperCase()] || null;
        }
        return { host: host, platform: platform };
    }
    catch (error) {
        return { host: 'WEB', platform: null };
    }
    var _b;
}
/** Convert the Office.context.host value to one of the Office JS Helpers constants. */
function convertHostValue(host) {
    var officeJsToHelperEnumMapping = {
        'Word': exports.HostType.WORD,
        'Excel': exports.HostType.EXCEL,
        'PowerPoint': exports.HostType.POWERPOINT,
        'Outlook': exports.HostType.OUTLOOK,
        'OneNote': exports.HostType.ONENOTE,
        'Project': exports.HostType.PROJECT,
        'Access': exports.HostType.ACCESS
    };
    return officeJsToHelperEnumMapping[host] || null;
}
/** Convert the Office.context.platform value to one of the Office JS Helpers constants. */
function convertPlatformValue(platform) {
    var officeJsToHelperEnumMapping = {
        'PC': exports.PlatformType.PC,
        'OfficeOnline': exports.PlatformType.OFFICE_ONLINE,
        'Mac': exports.PlatformType.MAC,
        'iOS': exports.PlatformType.IOS
    };
    return officeJsToHelperEnumMapping[platform] || null;
}
/**
 * Helper exposing useful Utilities for Office-Add-ins.
 */
var Utilities = (function () {
    function Utilities() {
    }
    Object.defineProperty(Utilities, "host", {
        /*
         * Returns the current host which is either the name of the application where the
         * Office Add-in is running ("EXCEL", "WORD", etc.) or simply "WEB" for all other platforms.
         * The property is always returned in ALL_CAPS.
         * Note that this property is guaranteed to return the correct value ONLY after Office has
         * initialized (i.e., inside, or sequentially after, an Office.initialize = function() { ... }; statement).
         *
         * This code currently uses a workaround that relies on the internals of Office.js.
         * A more robust approach is forthcoming within the official  Office.js library.
         * Once the new approach is released, this implementation will switch to using it
         * instead of the current workaround.
         */
        get: function () {
            return getHostInfo().host;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Utilities, "platform", {
        /*
        * Returns the host application's platform ("IOS", "MAC", "OFFICE_ONLINE", or "PC").
        * This is only valid for Office Add-ins, and hence returns null if the HostType is WEB.
        * The platform is in ALL-CAPS.
        * Note that this property is guaranteed to return the correct value ONLY after Office has
        * initialized (i.e., inside, or sequentially after, an Office.initialize = function() { ... }; statement).
        *
        * This code currently uses a workaround that relies on the internals of Office.js.
        * A more robust approach is forthcoming within the official  Office.js library.
        * Once the new approach is released, this implementation will switch to using it
        * instead of the current workaround.
        */
        get: function () {
            return getHostInfo().platform;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Utilities, "isAddin", {
        /**
         * Utility to check if the code is running inside of an add-in.
         */
        get: function () {
            return Utilities.host !== exports.HostType.WEB;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Utility to print prettified errors.
     * If multiple parameters are sent then it just logs them instead.
     */
    Utilities.log = function (exception, extras) {
        var args = [];
        for (var _i = 2; _i < arguments.length; _i++) {
            args[_i - 2] = arguments[_i];
        }
        if (!(extras == null)) {
            return console.log.apply(console, [exception, extras].concat(args));
        }
        if (exception == null) {
            console.error(exception);
        }
        else if (typeof exception === 'string') {
            console.error(exception);
        }
        else {
            console.group(exception.name + ": " + exception.message);
            {
                var innerException = exception;
                if (exception instanceof custom_error_1.CustomError) {
                    innerException = exception.innerError;
                }
                if (window.OfficeExtension && innerException instanceof OfficeExtension.Error) {
                    console.groupCollapsed('Debug Info');
                    console.error(innerException.debugInfo);
                    console.groupEnd();
                }
                {
                    console.groupCollapsed('Stack Trace');
                    console.error(exception.stack);
                    console.groupEnd();
                }
                {
                    console.groupCollapsed('Inner Error');
                    console.error(innerException);
                    console.groupEnd();
                }
            }
            console.groupEnd();
        }
    };
    return Utilities;
}());
exports.Utilities = Utilities;
//# sourceMappingURL=utilities.js.map