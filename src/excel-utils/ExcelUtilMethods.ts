import * as XLSX from "xlsx";

export function handleMergedCells(sheet: XLSX.WorkSheet): XLSX.WorkSheet {
  if (!sheet["!merges"]) return sheet; // No merged cells to process

  // Loop through each merge range and fill merged cells with the top-left cell's value
  sheet["!merges"].forEach((merge) => {
    const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const startCellValue = sheet[startCell]?.v;

    // Loop over each cell in the merged range and set the value to the top-left cell's value
    for (let R = merge.s.r; R <= merge.e.r; ++R) {
      for (let C = merge.s.c; C <= merge.e.c; ++C) {
        const cell = XLSX.utils.encode_cell({ r: R, c: C });
        sheet[cell] = { v: startCellValue !== undefined ? startCellValue : "" }; // Fill with the top-left cell's value
      }
    }
  });

  // Trim whitespace from string values and ensure all rows are of equal length
  const range = XLSX.utils.decode_range(sheet["!ref"]!);
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = XLSX.utils.encode_cell({ r: R, c: C });
      const cellValue = sheet[cell]?.v;

      // Trim whitespace for strings
      if (typeof cellValue === "string") {
        sheet[cell].v = cellValue.trim();
      }

      // Ensure cell exists and fill with an empty string if missing
      if (!sheet[cell]) {
        sheet[cell] = { v: "" }; // Fill empty cells with an empty string
      }
    }
  }

  return sheet;
}
