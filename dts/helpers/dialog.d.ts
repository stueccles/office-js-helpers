import { CustomError } from '../errors/custom.error';
/**
 * Custom error type to handle API specific errors.
 */
export declare class DialogError extends CustomError {
    innerError: Error;
    /**
     * @constructor
     *
     * @param message Error message to be propagated.
     * @param state OAuth state if available.
    */
    constructor(message: string, innerError?: Error);
}
/**
 * An optimized size object computed based on Screen Height & Screen Width
 */
export interface IDialogSize {
    /**
     * Width in pixels
     */
    width: number;
    /**
     * Width in percentage
     */
    width$: number;
    /**
     * Height in pixels
     */
    height: number;
    /**
     * Height in percentage
     */
    height$: number;
}
export declare class Dialog<T> {
    url: string;
    useTeamsDialog: boolean;
    /**
     * @constructor
     *
     * @param url Url to be opened in the dialog.
     * @param width Width of the dialog.
     * @param height Height of the dialog.
    */
    constructor(url?: string, width?: number, height?: number, useTeamsDialog?: boolean);
    private _result;
    readonly result: Promise<T>;
    size: IDialogSize;
    private _addinDialog();
    private _teamsDialog();
    /**
     * Close any open dialog by providing an optional message.
     * If more than one dialogs are attempted to be opened
     * an expcetion will be created.
     */
    static close(message?: any, useTeamsDialog?: boolean): void;
    private _optimizeSize(width, height);
    private _maxSize(value, max);
    private _percentage(value, max);
    private _safeParse(data);
}
