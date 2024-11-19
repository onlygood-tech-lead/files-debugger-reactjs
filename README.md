# Summary

Project to debug input files and improve testing

# Build Project

- Clone the repo
- Install all dependencies using `npm install`
- Start the development environment using `npm run start`

# Excel Helpers

- `ExcelUtilMethods`: This file contains methods for processing an excel worksheet to handle merged cells, trim whitespaces, enusre consistent row length and replace any null values with the appropriate merged values. ***handleMergedCells()*** is called directly on the instance of *XLSX.Worksheet*, followed by ***worksheetToArray()*** to populate merged cell data.
