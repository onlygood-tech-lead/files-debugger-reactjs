import { ExcelDateUtils } from "../excel-utils/ExcelDateUtils";

// Define the structure for a single data point
export type IotDataPoint = {
  plant: string;
  date: Date;
  hourRange: string;
  sensorName: string;
  consumedKW: number;
};

// Define possible error types
enum ParseErrorType {
  InvalidDate,
  InvalidConsumption,
  MissingData,
  InvalidData,
}

// Private custom error class
class ParseError extends Error {
  constructor(
    public row: number,
    public type: ParseErrorType,
    message: string,
  ) {
    super(message);
    this.name = "ParseError";
  }
}

class IotAnalyticsMetalmanParser {
  private parsedData: IotDataPoint[] = [];

  // Main public method to parse the entire dataset
  public parseDataset(rawData: any[][]): {
    result: IotDataPoint[];
    errors?: ParseError[];
  } {
    const [_, ...rows] = rawData;
    const result: IotDataPoint[] = [];
    const errors: ParseError[] = [];

    rows.forEach((row, index) => {
      try {
        const dataPoint = this.parseRow(row, index + 1);
        if (dataPoint) {
          result.push(dataPoint);
        }
      } catch (error) {
        if (error instanceof ParseError) {
          errors.push(error);
        }
      }
    });

    this.parsedData = result;
    return errors.length > 0 ? { result, errors } : { result };
  }

  // Private helper method to parse a single row
  private parseRow(row: any[], rowIndex: number): IotDataPoint | null {
    // Check if all values in the row are null, undefined, or empty strings
    if (
      row.every(
        (value) =>
          value == null || (typeof value === "string" && value.trim() === ""),
      )
    ) {
      throw new ParseError(
        rowIndex,
        ParseErrorType.MissingData,
        "Row has insufficient data",
      );
    }

    const [plant, dateValue, hourRange, sensorName, consumedKW] = row;

    // Validate and convert date using ExcelDateUtils
    if (!ExcelDateUtils.isValidExcelDate(dateValue)) {
      throw new ParseError(
        rowIndex,
        ParseErrorType.InvalidDate,
        `Invalid date value: ${dateValue}`,
      );
    }
    const date = ExcelDateUtils.toJsDate(dateValue);

    // Validate plant
    if (typeof plant !== "string" || plant.trim() === "") {
      throw new ParseError(
        rowIndex,
        ParseErrorType.InvalidData,
        "Invalid plant name",
      );
    }

    // Validate hourRange
    if (typeof hourRange !== "string" || hourRange.trim() === "") {
      throw new ParseError(
        rowIndex,
        ParseErrorType.InvalidData,
        "Invalid hour range",
      );
    }

    // Validate sensorName
    if (typeof sensorName !== "string" || sensorName.trim() === "") {
      throw new ParseError(
        rowIndex,
        ParseErrorType.InvalidData,
        "Invalid sensor name",
      );
    }

    // Validate consumedKW
    const parsedConsumedKW = Number(consumedKW);
    if (isNaN(parsedConsumedKW)) {
      throw new ParseError(
        rowIndex,
        ParseErrorType.InvalidConsumption,
        `Invalid consumedKW value: ${consumedKW}`,
      );
    }

    return {
      plant: plant.trim(),
      date,
      hourRange: hourRange.trim(),
      sensorName: sensorName.trim(),
      consumedKW: parsedConsumedKW,
    };
  }

  /**
   * Get all unique sensor names from the parsed data.
   * @returns An array of unique sensor names.
   */
  public getUniqueSensorNames(): string[] {
    const sensorNames = new Set(
      this.parsedData.map((point) => point.sensorName),
    );
    return Array.from(sensorNames);
  }

  /**
   * Get the latest consumed value for a given sensor name.
   * @param sensorName The name of the sensor to look up.
   * @returns The latest consumed value, date, and hour range for the given sensor, or null if not found.
   */
  public getLatestConsumedValueForSensor(
    sensorName: string,
  ): { value: number; date: Date; hourRange: string } | null {
    const sensorData = this.parsedData.filter(
      (point) => point.sensorName === sensorName,
    );
    if (sensorData.length === 0) {
      return null;
    }

    // Assuming the dataset is already sorted by date and time in descending order
    const latestData = sensorData[sensorData.length - 1];

    return {
      value: latestData.consumedKW,
      date: latestData.date,
      hourRange: latestData.hourRange,
    };
  }

  /**
   * Calculate the total consumption for a given list of IotDataPoints.
   * @param dataPoints The list of IotDataPoints to sum.
   * @returns The total consumed KW.
   */
  private calculateTotalConsumption(dataPoints: IotDataPoint[]): number {
    return dataPoints.reduce((total, point) => total + point.consumedKW, 0);
  }

  /**
   * Get the total consumption based on optional filters.
   * @param options Optional filters for the consumption calculation.
   * @returns The total consumed KW based on the provided filters.
   */
  public getTotalConsumption(options?: {
    plant?: string;
    sensorName?: string;
    startDate?: Date;
    endDate?: Date;
  }): number {
    let filteredData = this.parsedData;

    if (options) {
      if (options.plant) {
        filteredData = filteredData.filter(
          (point) => point.plant === options.plant,
        );
      }

      if (options.sensorName) {
        filteredData = filteredData.filter(
          (point) => point.sensorName === options.sensorName,
        );
      }

      if (options.startDate) {
        filteredData = filteredData.filter(
          (point) => point.date >= options.startDate!,
        );
      }

      if (options.endDate) {
        filteredData = filteredData.filter(
          (point) => point.date <= options.endDate!,
        );
      }
    }

    return this.calculateTotalConsumption(filteredData);
  }

  /**
   * Get the total consumption for a specific plant.
   * @param plant The name of the plant.
   * @returns The total consumed KW for the specified plant.
   */
  public getTotalConsumptionForPlant(plant: string): number {
    return this.getTotalConsumption({ plant });
  }

  /**
   * Get the total consumption for a specific sensor.
   * @param sensorName The name of the sensor.
   * @returns The total consumed KW for the specified sensor.
   */
  public getTotalConsumptionForSensorName(sensorName: string): number {
    return this.getTotalConsumption({ sensorName });
  }

  /**
   * Get the total consumption for a specific sensor within a date range.
   * @param sensorName The name of the sensor.
   * @param startDate The start date of the range.
   * @param endDate The end date of the range.
   * @returns The total consumed KW for the specified sensor within the date range.
   */
  public getTotalConsumptionForSensorNameInDateRange(
    sensorName: string,
    startDate: Date,
    endDate: Date,
  ): number {
    return this.getTotalConsumption({ sensorName, startDate, endDate });
  }

  /**
   * Clears the parsed dataset and performs memory cleanup.
   */
  public clearDataset(): void {
    // Clear the parsed data array
    this.parsedData = [];

    // Force garbage collection (if available)
    if (global.gc) {
      global.gc();
    }
  }
}

// Utility functions to consume the API
/**
 * Parses the given IoT dataset.
 * @param rawData The raw data from the Excel file.
 * @returns An object containing the parsed results and any errors encountered.
 */
export function parseIotDataset(rawData: any[][]): {
  result: IotDataPoint[];
  errors?: ParseError[];
} {
  const parser = new IotAnalyticsMetalmanParser();
  return parser.parseDataset(rawData);
}

// Export types and enums for consumers of the API
export { ParseError, ParseErrorType };

// Export the class (optional, if you want to allow direct usage)
export { IotAnalyticsMetalmanParser };
