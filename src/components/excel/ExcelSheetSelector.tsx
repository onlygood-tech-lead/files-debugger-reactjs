type ExcelSheetSelectorProps = {
  currentFile: File;
  sheetNames: string[];
  selectedSheetName: string;
  setSelectedSheet: (name: string) => void;
};

export default function ExcelSheetSelector({
  currentFile,
  sheetNames,
  selectedSheetName,
  setSelectedSheet,
}: ExcelSheetSelectorProps) {
  return (
    <div className="overflow-auto max-w-full">
      <div className="flex justify-start items-start gap-2 mb-2">
        <h3 className="text-lg font-normal">Select a Worksheet:</h3>
        <h4 className="text-lg font-semibold italic">{currentFile.name}</h4>
      </div>
      <ul className="flex space-x-4 overflow-x-auto p-2 border rounded">
        {sheetNames.map((sheetName) => (
          <li key={sheetName}>
            <button
              className={`px-4 py-2 rounded ${
                selectedSheetName === sheetName
                  ? "bg-blue-500 text-white"
                  : "bg-gray-200"
              }`}
              onClick={() => setSelectedSheet(sheetName)}
            >
              {sheetName}
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
