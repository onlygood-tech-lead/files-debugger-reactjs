import * as XLSX from "xlsx";

// interface required for typescript
interface ProgressEvent<T extends EventTarget = EventTarget> extends Event {
  readonly lengthComputable: boolean;
  readonly loaded: number;
  readonly target: T | null;
  readonly total: number;
}

export function readContentsFromExcelFile(file: File): Promise<XLSX.WorkBook> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (event: ProgressEvent<FileReader>) => {
      try {
        if (event.target && event.target.result instanceof ArrayBuffer) {
          const data = event.target.result;
          const workbook = XLSX.read(data, { type: "array" });
          resolve(workbook);
        } else {
          throw new Error(`Invalid file data`);
        }
      } catch (error) {
        reject(error);
      } finally {
        cleanupReader(reader);
      }
    };

    reader.onerror = () => {
      reject(new Error(`File reader failed`));
      cleanupReader(reader);
    };

    reader.onabort = () => {
      reject(new Error(`File reader aborted`));
      cleanupReader(reader);
    };

    // call the main method to read file as array buffer
    try {
      reader.readAsArrayBuffer(file);
    } catch (error) {
      reject(error);
      cleanupReader(reader);
    }
  });
}

function cleanupReader(reader: FileReader): void {
  if (reader) {
    reader.onload = null;
    reader.onerror = null;
    reader.onabort = null;
  }
}
