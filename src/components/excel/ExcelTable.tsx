import * as XLSX from "xlsx";

type ExcelTableProps = {
  worksheetData: XLSX.WorkSheet;
};

export default function ExcelTable({ worksheetData }: ExcelTableProps) {
  const sheet = worksheetData;
  const data: any[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  });

  // Retrieve the merged cell ranges from the worksheet
  const merges = sheet["!merges"] || [];

  // Create a map to handle merged cell spans (rowSpan and colSpan)
  const mergeMap: Record<string, { rowSpan: number; colSpan: number }> = {};

  merges.forEach((merge) => {
    const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    const rowSpan = merge.e.r - merge.s.r + 1;
    const colSpan = merge.e.c - merge.s.c + 1;
    mergeMap[startCell] = { rowSpan, colSpan };

    // Fill other cells in the range with null to skip rendering
    for (let R = merge.s.r; R <= merge.e.r; R++) {
      for (let C = merge.s.c; C <= merge.s.c; C++) {
        if (!(R === merge.s.r && C === merge.s.c)) {
          //   const cell = XLSX.utils.encode_cell({ r: R, c: C });
          data[R][C] = null; // Set cells within the merge range to null
        }
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
                if (cell === null) return null; // Skip cells that are part of a merged range

                const cellKey = XLSX.utils.encode_cell({
                  r: rowIndex,
                  c: colIndex,
                });
                const merge = mergeMap[cellKey];

                return (
                  <td
                    key={colIndex}
                    className="border border-gray-300 p-2"
                    rowSpan={merge?.rowSpan || 1}
                    colSpan={merge?.colSpan || 1}
                  >
                    {cell !== undefined ? cell : ""}
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
