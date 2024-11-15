import React, { useState, useMemo } from "react";
import { readContentsFromExcelFile } from "../../excel-utils/ExcelFileReader";
import { AiOutlineFileExcel, AiOutlineDelete } from "react-icons/ai";
import * as XLSX from "xlsx";
import ExcelSheetSelector from "./ExcelSheetSelector";
import JsonViewer from "../json-viewer/JsonViewer";
import {
  handleMergedCells,
  worksheetToArray,
} from "../../excel-utils/ExcelUtilMethods";
import ExcelTable from "./ExcelTable";

export default function ExcelReaderView() {
  const [workbookData, setWorkbookData] = useState<XLSX.WorkBook | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState<string>("");
  const [useCustomViewer, setUseCustomViewer] = useState<boolean>(true);

  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    try {
      const file = event.target.files?.[0] as File;
      const workbook = await readContentsFromExcelFile(file);
      setWorkbookData(workbook);
    } catch (err) {
      console.error("Error parsing file:", err);
    }
  };

  const handleReset = () => {
    setWorkbookData(null);
    setSelectedSheetName("");
  };

  // Memoized value for the currently selected sheet's data
  const selectedSheetData = useMemo(() => {
    if (!workbookData || !selectedSheetName) return null;
    // Get the sheet and handle merged cells before converting to JSON
    const sheet = workbookData.Sheets[selectedSheetName];
    const processedSheet = handleMergedCells(sheet); // Process merged cells
    return worksheetToArray(processedSheet);
  }, [workbookData, selectedSheetName]);

  const selectedSheet = workbookData
    ? workbookData.Sheets[selectedSheetName]
    : null;

  return (
    <div className="p-6 space-y-4">
      <div className="flex items-center space-x-4">
        {!workbookData ? (
          <label className="bg-blue-500 text-white px-4 py-2 rounded cursor-pointer hover:bg-blue-600">
            <AiOutlineFileExcel className="inline-block mr-2" />
            Upload Excel
            <input
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={handleFileUpload}
            />
          </label>
        ) : (
          <button
            className="bg-red-500 text-white px-4 py-2 rounded hover:bg-red-600"
            onClick={handleReset}
          >
            <AiOutlineDelete className="inline-block mr-2" />
            Reset
          </button>
        )}
      </div>

      {workbookData && (
        <div className="space-y-12">
          <ExcelSheetSelector
            sheetNames={workbookData.SheetNames}
            selectedSheetName={selectedSheetName}
            setSelectedSheet={setSelectedSheetName}
          />

          {/* {selectedSheet && (
            <pre className="bg-gray-100 p-4 rounded border overflow-auto max-w-full">
              {JSON.stringify(selectedSheet, null, 2)}
            </pre>
          )} */}

          {selectedSheetData && (
            <div className="space-y-4">
              {selectedSheet && <ExcelTable worksheetData={selectedSheet} />}
              <div className="flex items-center justify-between gap-4">
                <h4 className="text-lg font-semibold">Sheet Data:</h4>
                <button
                  className="bg-gray-500 text-white px-4 py-2 rounded hover:bg-gray-600"
                  onClick={() => setUseCustomViewer(!useCustomViewer)} // Toggle viewer
                >
                  {useCustomViewer
                    ? "Switch to Default View"
                    : "Switch to Custom Viewer"}
                </button>
              </div>
              {useCustomViewer ? (
                <JsonViewer
                  src={selectedSheetData}
                  theme="monokai"
                  collapsed={false}
                  enableClipboard={true}
                  displayDataTypes={true}
                  indentWidth={2}
                  collapseStringsAfterLength={50}
                />
              ) : (
                <pre className="bg-gray-100 p-4 rounded border overflow-auto max-w-full">
                  {JSON.stringify(selectedSheetData, null, 2)}
                </pre>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
