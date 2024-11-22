import { WorkBook } from "xlsx";
import { ExcelDateUtils } from "../excel-utils/ExcelDateUtils";
import {
  handleMergedCells,
  worksheetToArray,
} from "../excel-utils/ExcelUtilMethods";
import { matchNormalizedPhrase } from "../utils/general.helpers";

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

class IotDataQuery {
  private filters: {
    plant?: string;
    sensorName?: string;
    startDate?: Date;
    endDate?: Date;
  } = {};

  /**
   * Creates an instance of IotDataQuery.
   * @param data An array of IotDataPoint objects to query.
   */
  constructor(private data: IotDataPoint[]) {}

  /**
   * Filters the data for a specific plant.
   * @param plantName The name of the plant to filter for.
   * @returns The IotDataQuery instance for method chaining.
   */
  public forPlant(plantName: string): IotDataQuery {
    this.filters.plant = plantName;
    return this;
  }

  /**
   * Filters the data for a specific plant and sensor.
   * @param plantName The name of the plant to filter for.
   * @param sensorName The name of the sensor to filter for.
   * @returns The IotDataQuery instance for method chaining.
   */
  public forSensor(plantName: string, sensorName: string): IotDataQuery {
    this.filters.plant = plantName;
    this.filters.sensorName = sensorName;
    return this;
  }

  /**
   * Filters the data for a specific date range.
   * @param startDate The start date of the range.
   * @param endDate The end date of the range.
   * @returns The IotDataQuery instance for method chaining.
   */
  public between(startDate: Date, endDate: Date): IotDataQuery {
    this.filters.startDate = startDate;
    this.filters.endDate = endDate;
    return this;
  }

  /**
   * Calculates the sum of consumedKW values for the filtered data.
   * @returns The total sum of consumedKW values.
   */
  public sumConsumedKW(): number {
    return this.applyFilters().reduce(
      (sum, point) => sum + point.consumedKW,
      0,
    );
  }

  /**
   * Calculates daily sums of consumedKW values for the filtered data.
   * @returns An array of objects containing the date and summed value for each day.
   */
  public dailySums(): { date: string; summedValue: number }[] {
    const dailySums = new Map<string, number>();
    this.applyFilters().forEach((point) => {
      const dateKey = point.date.toISOString().split("T")[0];
      const currentSum = dailySums.get(dateKey) || 0;
      dailySums.set(dateKey, currentSum + point.consumedKW);
    });
    return Array.from(dailySums.entries()).map(([date, summedValue]) => ({
      date,
      summedValue,
    }));
  }

  /**
   * Calculates daily sums of consumedKW values for each sensor in a specific plant.
   * @param plantName The name of the plant to calculate sums for.
   * @returns An array of objects containing the date and summed values for each sensor.
   */
  public dailySumsBySensor(plantName: string): {
    date: string;
    sensorSums: { [sensorName: string]: number };
  }[] {
    const dailySums = new Map<string, Map<string, number>>();

    this.applyFilters()
      .filter((point) => point.plant === plantName)
      .forEach((point) => {
        const dateKey = point.date.toISOString().split("T")[0];
        if (!dailySums.has(dateKey)) {
          dailySums.set(dateKey, new Map<string, number>());
        }
        const sensorSums = dailySums.get(dateKey)!;
        const currentSum = sensorSums.get(point.sensorName) || 0;
        sensorSums.set(point.sensorName, currentSum + point.consumedKW);
      });

    return Array.from(dailySums.entries()).map(([date, sensorSums]) => ({
      date,
      sensorSums: Array.from(sensorSums.entries()).reduce(
        (obj, [sensor, value]) => {
          obj[sensor] = value;
          return obj;
        },
        {} as { [sensorName: string]: number },
      ),
    }));
  }

  /**
   * Calculates the sum of consumedKW values for a specific plant and sensor.
   * @param plantName The name of the plant.
   * @param sensorName The name of the sensor.
   * @returns The total sum of consumedKW values for the specified plant and sensor.
   */
  public sumConsumedKWForSensor(plantName: string, sensorName: string): number {
    return this.data
      .filter(
        (point) => point.plant === plantName && point.sensorName === sensorName,
      )
      .reduce((sum, point) => sum + point.consumedKW, 0);
  }

  /**
   * Applies the current filters to the data.
   * @returns An array of IotDataPoint objects that match the current filters.
   */
  public applyFilters(): IotDataPoint[] {
    return this.data.filter((point) => {
      return (
        (!this.filters.plant || point.plant === this.filters.plant) &&
        (!this.filters.sensorName ||
          point.sensorName === this.filters.sensorName) &&
        (!this.filters.startDate || point.date >= this.filters.startDate) &&
        (!this.filters.endDate || point.date <= this.filters.endDate)
      );
    });
  }
}

class IotAnalyticsMetalmanParser {
  private parsedData: IotDataPoint[] = [];
  private readonly MAIN_WORKSHEET_NAME: string = "Master_Data";

  public query(): IotDataQuery {
    return new IotDataQuery(this.parsedData);
  }

  // method to parse worksheet name from array
  public parseMasterSheetFromWorkbook(workbook: WorkBook): any[][] {
    // determine whether worksheet is found
    const matchedSheetName = matchNormalizedPhrase(
      this.MAIN_WORKSHEET_NAME,
      workbook.SheetNames,
    )!;
    if (!workbook.SheetNames.includes(matchedSheetName)) {
      throw new Error(
        `Unable to find ${matchedSheetName} worksheet in workbook`,
      );
    }
    // parse this worksheet
    const processedSheet = handleMergedCells(workbook.Sheets[matchedSheetName]);
    return worksheetToArray(processedSheet);
  }

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

    this.parsedData.push(...result);
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
   * Get all unique plant names from the parsed data.
   * @returns An array of unique plant names.
   */
  public getUniquePlantNames(): string[] {
    return Array.from(
      new Set(
        this.query()
          .forPlant("")
          .applyFilters()
          .map((point) => point.plant),
      ),
    );
  }

  /**
   * Get all unique sensor names for a specific plant from the parsed data.
   * @param plantName The name of the plant to filter sensors for.
   * @returns An array of unique sensor names for the given plant.
   */
  public getUniqueSensorNamesForPlant(plantName: string): string[] {
    return Array.from(
      new Set(
        this.query()
          .forPlant(plantName)
          .applyFilters()
          .map((point) => point.sensorName),
      ),
    );
  }

  /**
   * Get the latest consumed value for a given sensor name in a specific plant.
   * @param plantName The name of the plant to look up.
   * @param sensorName The name of the sensor to look up.
   * @returns The latest consumed value, date, and hour range for the given sensor in the specified plant, or null if not found.
   */
  public getLatestConsumedValueForSensor(
    plantName: string,
    sensorName: string,
  ): { value: number; date: Date; hourRange: string } | null {
    const latestData = this.query()
      .forSensor(plantName, sensorName)
      .applyFilters()
      .sort((a, b) => b.date.getTime() - a.date.getTime())[0];

    return latestData
      ? {
          value: latestData.consumedKW,
          date: latestData.date,
          hourRange: latestData.hourRange,
        }
      : null;
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
