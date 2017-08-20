/// <reference types="office-js" />
/**
 * Helper exposing useful Utilities for Excel Add-ins.
 */
export declare class ExcelUtilities {
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
    static forceCreateSheet(workbook: Excel.Workbook, sheetName: string, clearOnly?: boolean): Promise<Excel.Worksheet>;
}
