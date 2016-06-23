# Opuxl - Opus Excel Addin

# Testing/Debugging

- OpuxlClassLibrary/Properties
- Debug
- Start External Project: <Your EXCEL.exe>
- Command Line Arguments: <Your build XLL File>
(eg: "C:\Users\stepan\Documents\Visual Studio 2015\Projects\Opuxl\OpuxlClassLibrary\bin\Debug\Opuxl.xll")
- Start with Debug Settings

# Packaging for distribution

- Add Pack="true" to the ExternalLibrary property of your Opuxl.dna configuration file
- Rebuild the project as a release build
- Run the ExcelDnaPack script in your terminal
(eg: "ExcelDna\Distribution\ExcelDnaPack.exe OpuxlClassLibrary\bin\Release\Opuxl.dna")
- Result is a packed .xll file