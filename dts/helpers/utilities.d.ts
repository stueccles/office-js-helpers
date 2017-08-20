import { CustomError } from '../errors/custom.error';
/**
 * Constant strings for the host types
 */
export declare const HostType: {
    WEB: string;
    ACCESS: string;
    EXCEL: string;
    ONENOTE: string;
    OUTLOOK: string;
    POWERPOINT: string;
    PROJECT: string;
    WORD: string;
};
/**
 * Constant strings for the host platforms
 */
export declare const PlatformType: {
    IOS: string;
    MAC: string;
    OFFICE_ONLINE: string;
    PC: string;
};
/**
 * Helper exposing useful Utilities for Office-Add-ins.
 */
export declare class Utilities {
    static readonly host: string;
    static readonly platform: string;
    /**
     * Utility to check if the code is running inside of an add-in.
     */
    static readonly isAddin: boolean;
    /**
     * Utility to print prettified errors.
     * If multiple parameters are sent then it just logs them instead.
     */
    static log(exception: Error | CustomError | string, extras?: any, ...args: any[]): void;
}
