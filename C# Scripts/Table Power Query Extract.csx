#r "C:\Program Files\Tabular Editor 3\Microsoft.Office.Interop.Excel.dll" // <-- copy dll from "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel.dll" and paste on this location


// Reference: Microsoft.Office.Interop.Excel is required for this script to work.
// Ensure you have Excel installed and accessible from your system.
using Microsoft.Office.Interop.Excel;
using System.IO;
using System;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Converters;

// Define the Excel file path
string excelFilePath = Environment.ExpandEnvironmentVariables(@"%userprofile%\Documents\Table_Source.xlsx"); // <-- path of file to be corrected

// Extract parameter value (e.g., #"Server Name")
string GrabValue(string mCode, string pattern, int patternIndex)
{
    var result = "";
    var match = Regex.Match(mCode, pattern);

    if (match.Success)
    {
        
        var arg = match.Groups[patternIndex].Value.Trim();
        Output(arg);
        var parameterName = "";
        
        // case: "Argument"
        if( arg.StartsWith("\"") ) {
            result = arg;
        } 
        
        // case: #"Argument"
        else if( arg.StartsWith("#") ) {
            parameterName = arg.Substring(2, arg.Length - 3);
            result =  Model.Expressions.FirstOrDefault(p => p.Name == parameterName).Expression;
        } 

        // case: Argument
        else {
            parameterName = arg;
            result =  Model.Expressions.FirstOrDefault(p => p.Name == parameterName).Expression;
        }

        
        match = Regex.Match( result, """([^""]+)""" );
        if (match.Success) 
            result = match.Groups[1].Value;
    }
    return result;

}


// Initialize Excel application
var excelApp = new Application();
excelApp.Visible = false;
excelApp.DisplayAlerts = false;

Workbook workbook;
Worksheet worksheet;

// Check if the file exists
if (File.Exists(excelFilePath))
{
    workbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly: false);
    worksheet = workbook.Sheets[1];
}
else
{
    workbook = excelApp.Workbooks.Add();
    worksheet = workbook.Sheets[1];
    worksheet.Cells[1, 1].Value = "Dataset";
    worksheet.Cells[1, 2].Value = "Table";
    worksheet.Cells[1, 3].Value = "Server";
    worksheet.Cells[1, 4].Value = "DB";
    worksheet.Cells[1, 5].Value = "Schema";
    worksheet.Cells[1, 6].Value = "DB View/Table";
}

// Start appending data from the last row
int lastRow = worksheet.Cells[worksheet.Rows.Count, 1].End(XlDirection.xlUp).Row + 1;
string Dataset = Model.Database.Name;
string PQry = "";
string vwName;

// Loop through all tables in the model
foreach (var table in Model.Tables)
{
    string sourceType = $"{table.SourceType}";
    if( sourceType == "M" ||sourceType == "PolicyRange")
    {
        if( sourceType == "PolicyRange" )
        {
            PQry = table.SourceExpression; // Extracting power query
        }
        else
        {
            PQry = table.Partitions[0].Expression;
        }

        if( PQry?.Contains("Sql.Database") == true || PQry?.Contains("PostGProd") == true )
        {
            worksheet.Cells[lastRow, 1].Value = Dataset;
            worksheet.Cells[lastRow, 2].Value = table.Name; // Table Name 
             
            worksheet.Cells[lastRow, 3].Value = GrabValue(PQry,  @"Sql\.Database\s*\(\s*([^,\)]+)\s*,\s*([^,\)]+)", 1 );
            worksheet.Cells[lastRow, 4].Value = GrabValue(PQry,  @"Sql\.Database\s*\(\s*([^,\)]+)\s*,\s*([^,\)]+)", 2 );
            worksheet.Cells[lastRow, 5].Value = GrabValue(PQry,  @"Schema\s*=\s*(#?""[^""]+""|\w+)\s*,\s*Item\s*=\s*(#?""[^""]+""|\w+)", 1 );
            worksheet.Cells[lastRow, 6].Value = GrabValue(PQry,   @"Schema\s*=\s*(#?""[^""]+""|\w+)\s*,\s*Item\s*=\s*(#?""[^""]+""|\w+)", 2 );
            lastRow++;
        
        }
        else{}
    } else{}
    PQry = null;
}

// Save and close the workbook

workbook.SaveAs(excelFilePath, XlFileFormat.xlOpenXMLWorkbook);
workbook.Close(false);
excelApp.Quit();

// Release COM objects
System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
worksheet = null;
workbook = null;
excelApp = null;
GC.Collect();
GC.WaitForPendingFinalizers();

Output("Export completed successfully!");