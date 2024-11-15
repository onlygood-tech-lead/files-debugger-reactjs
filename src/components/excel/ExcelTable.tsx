import * as XLSX from "xlsx";

type ExcelTableProps = {
  worksheetData: XLSX.WorkSheet;
};

export default function ExcelTable({ worksheetData }: ExcelTableProps) {
  // Convert worksheet data to a 2D array for easier manipulation
  const data: any[][] = XLSX.utils.sheet_to_json(worksheetData, {
    header: 1,
    defval: "",
  });

  // Retrieve merged cell information from the worksheet
  const merges = worksheetData["!merges"] || [];

  // Map to store merged cells with appropriate rowSpan and colSpan
  const mergeMap: Record<
    string,
    { rowSpan: number; colSpan: number; value: any }
  > = {};

  // Populate the mergeMap with rowSpan, colSpan, and top-left cell value for merged ranges
  merges.forEach((merge) => {
    const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const rowSpan = merge.e.r - merge.s.r + 1;
    const colSpan = merge.e.c - merge.s.c + 1;
    const topLeftValue = worksheetData[startCell]?.v || ""; // Use top-left cell value

    mergeMap[startCell] = { rowSpan, colSpan, value: topLeftValue };

    // Set all cells within the merged range to the top-left value
    for (let R = merge.s.r; R <= merge.e.r; R++) {
      for (let C = merge.s.c; C <= merge.e.c; C++) {
        if (R === merge.s.r && C === merge.s.c) {
          continue; // Skip the top-left cell as it's already set
        }
        const cellKey = XLSX.utils.encode_cell({ r: R, c: C });
        worksheetData[cellKey] = { v: topLeftValue }; // Ensure all merged cells have the same value
        data[R][C] = null; // Set other cells within the merge range to null to avoid duplicates
      }
    }
  });

  // Generate column headers (A, B, C, ...)
  const columnHeaders = Array.from({ length: data[0].length }, (_, i) =>
    String.fromCharCode(65 + i),
  );

  return (
    <div className="overflow-x-auto max-w-full">
      <table className="border-collapse border border-gray-400 min-w-[1200px]">
        <thead>
          <tr>
            <th className="border border-gray-300 p-2 bg-gray-200">#</th>
            {columnHeaders.map((header, index) => (
              <th
                key={index}
                className="border border-gray-300 p-2 bg-gray-200"
              >
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row, rowIndex) => (
            <tr key={rowIndex}>
              <td className="border border-gray-300 p-2 bg-gray-100">
                {rowIndex + 1}
              </td>
              {row.map((cell, colIndex) => {
                const cellKey = XLSX.utils.encode_cell({
                  r: rowIndex,
                  c: colIndex,
                });
                const merge = mergeMap[cellKey];

                // If the cell is part of a merged range and not the top-left, skip rendering
                if (cell === null) return null;

                return (
                  <td
                    key={colIndex}
                    className="border border-gray-300 p-2"
                    rowSpan={merge?.rowSpan || 1}
                    colSpan={merge?.colSpan || 1}
                  >
                    {merge ? merge.value : cell}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
