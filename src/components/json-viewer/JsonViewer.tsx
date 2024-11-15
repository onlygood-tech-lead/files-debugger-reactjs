import ReactJson, { ThemeKeys } from "react-json-view";

type JsonViewerProps = {
  src: object; // The JSON data to display
  theme?: ThemeKeys; // Theme for the viewer
  collapsed?: boolean | number; // True, false, or depth level to start collapsed
  enableClipboard?: boolean; // Enables copy-to-clipboard functionality
  displayDataTypes?: boolean; // Show/hide data type labels
  displayObjectSize?: boolean; // Show/hide size of the objects
  indentWidth?: number; // The indentation level
  collapsedDepth?: number; // Number of levels to collapse by default
  collapseStringsAfterLength?: number; // Collapse strings after a specified length
  name?: string | false; // Custom name for the root object, or false to hide
  iconStyle?: "circle" | "square" | "triangle"; // The icon style
  quotesOnKeys?: boolean; // Wrap keys with quotes in the viewer
  sortKeys?: boolean; // Sort keys alphabetically
  onEdit?: (edit: any) => void; // Callback for editing values
  onAdd?: (add: any) => void; // Callback for adding new values
  onDelete?: (del: any) => void; // Callback for deleting values
};

export default function JsonViewer({
  src,
  theme = "rjv-default",
  collapsed = false,
  enableClipboard = true,
  displayDataTypes = true,
  displayObjectSize = true,
  indentWidth = 4,
  collapseStringsAfterLength = 100,
  name = false,
  iconStyle = "triangle",
  quotesOnKeys = true,
  sortKeys = false,
  onEdit,
  onAdd,
  onDelete,
}: JsonViewerProps) {
  return (
    <ReactJson
      src={src}
      theme={theme}
      collapsed={collapsed}
      enableClipboard={enableClipboard}
      displayDataTypes={displayDataTypes}
      displayObjectSize={displayObjectSize}
      indentWidth={indentWidth}
      collapseStringsAfterLength={collapseStringsAfterLength}
      name={name}
      iconStyle={iconStyle}
      quotesOnKeys={quotesOnKeys}
      sortKeys={sortKeys}
      onEdit={onEdit}
      onAdd={onAdd}
      onDelete={onDelete}
    />
  );
}
