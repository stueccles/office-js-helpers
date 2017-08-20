import { CustomError } from './custom.error';
/**
 * Custom error type to handle API specific errors.
 */
export declare class APIError extends CustomError {
    innerError: Error;
    /**
     * @constructor
     *
     * @param message: Error message to be propagated.
     * @param innerError: Inner error if any
    */
    constructor(message: string, innerError?: Error);
}
