/**
 * Custom error type
 */
export declare abstract class CustomError extends Error {
    name: string;
    message: string;
    innerError: Error;
    constructor(name: string, message: string, innerError?: Error);
}
