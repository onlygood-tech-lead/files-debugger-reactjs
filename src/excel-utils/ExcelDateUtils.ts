export class ExcelDateUtils {
  private static readonly EXCEL_DATE_OFFSET = 25569;
  private static readonly MILLISECONDS_PER_DAY = 86400 * 1000;

  /**
   * Converts an Excel date number to a JavaScript Date object.
   * @param excelDate The Excel date number to convert.
   * @returns A JavaScript Date object representing the Excel date.
   */
  public static toJsDate(excelDate: number): Date {
    return new Date(
      (excelDate - this.EXCEL_DATE_OFFSET) * this.MILLISECONDS_PER_DAY,
    );
  }

  /**
   * Converts a JavaScript Date object to an Excel date number.
   * @param jsDate The JavaScript Date object to convert.
   * @returns The Excel date number representing the JavaScript Date.
   */
  public static toExcelDate(jsDate: Date): number {
    return (
      jsDate.getTime() / this.MILLISECONDS_PER_DAY + this.EXCEL_DATE_OFFSET
    );
  }

  /**
   * Checks if the given value is a valid Excel date number.
   * @param value The value to check.
   * @returns True if the value is a valid Excel date number, false otherwise.
   */
  public static isValidExcelDate(value: any): boolean {
    if (typeof value !== "number") return false;
    const date = this.toJsDate(value);
    return !isNaN(date.getTime());
  }
}
