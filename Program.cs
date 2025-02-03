using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using TechTalk.SpecFlow;
[Binding]
public class Program
{
    private List<List<string>> excelData;
    private List<List<string>> specFlowData;

    [Given(@"I verfiy merged cell ""([^""]*)""")]
    public void GivenIVerfiyMergedCell(string abcd)
    {
        var filePath = @"C:\Users\Anurag\Music\Compare.xlsx";  // Change this to your file path
        excelData = ReadExcelData(filePath);
        CompareMergedCells(excelData, abcd);
        
    }

    [Then(@"i verify below details")]
    public void ThenIVerifyBelowDetails(Table table)
    {
        specFlowData = ConvertTableToList(table);
        CompareNonMergedCells(excelData, specFlowData);
    }


    [Given(@"I verify headers ""([^""]*)"",""([^""]*)"",""([^""]*)"",""([^""]*)""")]
    public void GivenIVerifyHeaders(string p0, string p1, string p2, string p3)
    {
        var filePath = @"C:\Users\Anurag\Music\Compare.xlsx";
        IList<string> headers = new List<string>() { p0,p1,p2,p3 };
       IList<string> data= GetHeaderRow(filePath);
        for (int i = 0; i < data.Count; i++)
        {
            // Compare each header (including empty headers)
            Assert.AreEqual(data[i], headers[i], $"Header mismatch at position {i + 1}: Expected '{data[i]}', but got '{headers[i]}'");
        }

    }

    [Then(@"i verify below details only data")]
    public void ThenIVerifyBelowDetailsOnlyData(Table table)
    {
        var filePath = @"C:\Users\Anurag\Music\Compare.xlsx";  // Change this to your file path
        excelData = ReadExcelData(filePath);
        specFlowData = ConvertTableToList(table);
        CompareNonMergedCells(excelData, specFlowData);

    }



    public void GivenTheFollowingDataIsPresentInExcel()
    {
        var filePath = @"C:\Users\Anurag\Music\Compare.xlsx";  // Change this to your file path
        excelData = ReadExcelData(filePath);
      //  specFlowData = ConvertTableToList(table);
    }

    // Convert SpecFlow table to List<List<string>> for easier comparison
    //private List<List<string>> ConvertTableToList(Table table)
    //{
    //    var result = new List<List<string>>();
    //    foreach (var row in table.Rows)
    //    {
    //        var rowList = row.Values.ToList();
    //        result.Add(rowList);
    //    }
    //    return result;
    //}

    // Method to read data from Excel file using EPPlus
    private List<List<string>> ReadExcelData(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var data = new List<List<string>>();
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found at {filePath}");
        }

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Iterate through the rows and columns to read the data
            for (int row = 3; row <= rowCount; row++)
            {
                var rowData = new List<string>();

                for (int col = 1; col <= colCount; col++)
                {
                    string cellValue = string.Empty;

                    // Check if the cell is part of a merged range
                    var cell = worksheet.Cells[row, col];
                    if (cell.Merge)
                    {
                        // If merged, check if it's the first cell in the merged range
                        var mergedRange = worksheet.MergedCells.FirstOrDefault(r => r.Contains(cell.Address));

                        if (mergedRange != null && cell.Address == mergedRange.Split(':')[0])
                        {
                            // This is the first cell in the merged range, so get its value
                            cellValue = cell.Text.Trim();
                        }
                        else
                        {
                            // If it's not the first cell in the merged range, skip this cell entirely
                            continue;
                        }
                    }
                    else
                    {
                        // If the cell is not merged, get its value normally
                        cellValue = cell.Text.Trim();
                    }

                    // Add the value (only non-empty values and first merged cell value) to the row data
                    rowData.Add(cellValue);
                }

                // Add the row data to the list
                data.Add(rowData);
            }
        }

        return data;
    }

    //public static void Main(string[] args)
    //{
    //Program p= new Program();
    //    p.GivenTheFollowingDataIsPresentInExcel();
    //}

    // Convert SpecFlow table to List<List<string>> for easier comparison
    private List<List<string>> ConvertTableToList(Table table)
    {
        var result = new List<List<string>>();
        var headerRow = table.Header.ToList();  // Get headers
        result.Add(headerRow);
        foreach (var row in table.Rows)
        {
            var rowList = new List<string>();
            for (int i = 0; i < headerRow.Count; i++)
            {
                if (row.Values.Count > i)
                {
                    rowList.Add(row.Values.ElementAt(i));  // Add value from SpecFlow table
                }
                else
                {
                    rowList.Add(string.Empty);  // Add empty string for missing values
                }
            }
            result.Add(rowList);
        }
        return result;
    }
    // Separate method for comparing merged cells
    private void CompareMergedCells(List<List<string>> excelData, string data)
    {
        var excelValue = excelData[0][0];
        var specFlowValue = data;

        if (excelValue != specFlowValue)
        {
            Assert.Fail($"Data mismatch at merged cell row {0 + 1}, column {0 + 1}. Excel value: '{excelValue}', SpecFlow value: '{specFlowValue}'");
        }

        //// Loop through each row and column for the merged cells comparison
        //for (int i = 0; i < excelData.Count; i++)
        //{
        //    for (int j = 0; j < excelData[i].Count; j++)
        //    {
        //        // Check for merged cells (only compare the first cell of the merged range)
        //        var excelValue = excelData[i][j];
        //        var specFlowValue = data;

        //        if (excelValue != specFlowValue)
        //        {
        //            Assert.Fail($"Data mismatch at merged cell row {i + 1}, column {j + 1}. Excel value: '{excelValue}', SpecFlow value: '{specFlowValue}'");
        //        }
        //    }
        //}
    }
    // General comparison method (for non-merged cells)
    private void CompareNonMergedCells(List<List<string>> excelData, List<List<string>> specFlowData)
    {
        // Compare all non-merged cells as usual
        for (int i = 1; i < excelData.Count; i++)
        {
            for (int j = 0; j < excelData[i].Count; j++)
            {
                var excelValue = excelData[i][j];
                var specFlowValue = specFlowData[i][j];

                if (excelValue != specFlowValue)
                {
                    Assert.Fail($"Data mismatch at row {i + 1}, column {j + 1}. Excel value: '{excelValue}', SpecFlow value: '{specFlowValue}'");
                }
            }
        }
    }


    public static List<string> GetHeaderRow(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var headerRow = new List<string>();

        // Ensure the file exists
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found at {filePath}");
        }

        // Open the Excel file
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];  // Use the first worksheet
            var columnCount = worksheet.Dimension.Columns;   // Get the number of columns in the worksheet

            // Iterate through each column in the header row (first row)
            for (int col = 1; col <= columnCount; col++)
            {
                var headerCellValue = worksheet.Cells[2, col].Text.Trim();  // Retrieve value from the header cell
                headerRow.Add(headerCellValue);
            }
        }

        return headerRow;  // Return the header row as a list of strings
    }

}