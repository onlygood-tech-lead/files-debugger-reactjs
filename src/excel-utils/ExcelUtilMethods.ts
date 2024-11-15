import * as XLSX from "xlsx";

/**
 * Processes an Excel worksheet to handle merged cells, trim whitespace, ensure consistent row lengths,
 * and replace any null values with the appropriate merged values.
 * @param sheet The worksheet to process.
 * @returns The processed worksheet with merged cells handled, whitespace trimmed, consistent row lengths, and no null values.
 */
export function handleMergedCells(sheet: XLSX.WorkSheet): XLSX.WorkSheet {
  if (!sheet["!merges"]) return sheet; // No merged cells to process

  // Loop through each merge range and fill merged cells with the top-left cell's value
  sheet["!merges"].forEach((merge) => {
    const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const topLeftValue = sheet[startCell]?.v || ""; // Use top-left cell's value

    // Populate all cells within the merged range with the top-left cell value
    for (let R = merge.s.r; R <= merge.e.r; R++) {
      for (let C = merge.s.c; C <= merge.e.c; C++) {
        const cellKey = XLSX.utils.encode_cell({ r: R, c: C });
        sheet[cellKey] = { v: topLeftValue }; // Set all cells in merged range to top-left value
      }
    }
  });

  // Retrieve the range of cells to ensure consistent row lengths and handle missing cells
  const range = XLSX.utils.decode_range(sheet["!ref"]!);
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellKey = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = sheet[cellKey]?.v;

      // Trim whitespace for strings
      if (typeof cellValue === "string") {
        sheet[cellKey].v = cellValue.trim();
      }

      // Ensure cell exists and fill with an empty string if missing
      if (!sheet[cellKey]) {
        sheet[cellKey] = { v: "" }; // Fill empty cells with an empty string
      }
    }
  }

  return sheet;
}

/**
 * Converts the processed worksheet to a 2D array, ensuring that merged values are correctly filled in each cell of the merged range.
 * @param sheet The worksheet to convert.
 * @returns A 2D array representing the worksheet data, with merged cell values populated across ranges.
 */
export function worksheetToArray(sheet: XLSX.WorkSheet): any[][] {
  // Convert worksheet to a 2D array and ensure merged values are propagated across all cells in the merged range
  const data: any[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  });

  // Retrieve the merged cell information and propagate values
  const merges = sheet["!merges"] || [];
  merges.forEach((merge) => {
    const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const topLeftValue = sheet[startCell]?.v || ""; // Value of the top-left cell

    // Fill the 2D array for each cell in the merged range with the top-left value
    for (let R = merge.s.r; R <= merge.e.r; R++) {
      for (let C = merge.s.c; C <= merge.e.c; C++) {
        data[R][C] = topLeftValue; // Ensure the merged value is set in the array
      }
    }
  });

  return data;
}
