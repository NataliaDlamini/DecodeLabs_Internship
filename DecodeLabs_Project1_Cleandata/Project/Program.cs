using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace DataCleaningProject
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("====================================");
            Console.WriteLine("DATA CLEANING PROJECT - DECODELABS");
            Console.WriteLine("====================================\n");

            string inputFile = "Dataset for Data Analytics.xlsx";
            string outputFile = "Cleaned_Dataset_CSharp.xlsx";

            if (!File.Exists(inputFile))
            {
                Console.WriteLine("Dataset file not found!");
                return;
            }

            try
            {
                using (XLWorkbook workbook = new XLWorkbook(inputFile))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);

                    DataTable dt = new DataTable();

                    bool firstRow = true;

                    foreach (IXLRow row in worksheet.RowsUsed())
                    {
                        if (firstRow)
                        {
                            foreach (IXLCell cell in row.Cells())
                            {
                                dt.Columns.Add(cell.Value.ToString());
                            }

                            firstRow = false;
                        }
                        else
                        {
                            DataRow dataRow = dt.NewRow();

                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                dataRow[i] = row.Cell(i + 1).Value.ToString();
                            }

                            dt.Rows.Add(dataRow);
                        }
                    }

                    Console.WriteLine($"Rows Loaded: {dt.Rows.Count}");
                    Console.WriteLine($"Columns Loaded: {dt.Columns.Count}\n");

                    Console.WriteLine("Checking Missing Values...\n");

                    foreach (DataColumn column in dt.Columns)
                    {
                        int missingCount = 0;

                        foreach (DataRow row in dt.Rows)
                        {
                            if (string.IsNullOrWhiteSpace(row[column].ToString()))
                            {
                                missingCount++;
                                row[column] = "Unknown";
                            }
                        }

                        Console.WriteLine($"{column.ColumnName}: {missingCount} missing values");
                    }

                    Console.WriteLine("\nRemoving Duplicate Rows...\n");

                    var uniqueRows = dt.AsEnumerable()
                        .GroupBy(r => string.Join("|", r.ItemArray))
                        .Select(g => g.First());

                    DataTable cleanedTable = dt.Clone();

                    foreach (var row in uniqueRows)
                    {
                        cleanedTable.ImportRow(row);
                    }

                    Console.WriteLine($"Rows After Cleaning: {cleanedTable.Rows.Count}");

                    Console.WriteLine("\nChecking Duplicate Order IDs...\n");

                    if (cleanedTable.Columns.Contains("OrderID"))
                    {
                        var duplicateIDs = cleanedTable.AsEnumerable()
                            .GroupBy(r => r["OrderID"].ToString())
                            .Where(g => g.Count() > 1);

                        Console.WriteLine($"Duplicate Order IDs Found: {duplicateIDs.Count()}");

                        cleanedTable = cleanedTable.AsEnumerable()
                            .GroupBy(r => r["OrderID"].ToString())
                            .Select(g => g.First())
                            .CopyToDataTable();

                        Console.WriteLine("Duplicate Order IDs Removed!");
                    }

                    Console.WriteLine("\nCleaning Date Formats...\n");

                    if (cleanedTable.Columns.Contains("Date"))
                    {
                        foreach (DataRow row in cleanedTable.Rows)
                        {
                            DateTime parsedDate;

                            bool validDate = DateTime.TryParse(
                                row["Date"].ToString(),
                                out parsedDate
                            );

                            if (validDate)
                            {
                                row["Date"] = parsedDate.ToString("yyyy-MM-dd");
                            }
                        }

                        Console.WriteLine("Date Formatting Completed!");
                    }

                    Console.WriteLine("\nStandardizing Text Data...\n");

                    TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;

                    foreach (DataColumn column in cleanedTable.Columns)
                    {
                        foreach (DataRow row in cleanedTable.Rows)
                        {
                            string value = row[column].ToString();

                            value = value.Trim();

                            value = textInfo.ToTitleCase(value.ToLower());

                            row[column] = value;
                        }
                    }

                    Console.WriteLine("Text Formatting Completed!");

                    Console.WriteLine("\n====================================");
                    Console.WriteLine("FINAL QUALITY REPORT");
                    Console.WriteLine("====================================\n");

                    int remainingDuplicates = cleanedTable.AsEnumerable()
                        .GroupBy(r => string.Join("|", r.ItemArray))
                        .Count(g => g.Count() > 1);

                    Console.WriteLine($"Remaining Duplicate Rows: {remainingDuplicates}");

                    if (cleanedTable.Columns.Contains("OrderID"))
                    {
                        int remainingDuplicateIDs = cleanedTable.AsEnumerable()
                            .GroupBy(r => r["OrderID"].ToString())
                            .Count(g => g.Count() > 1);

                        Console.WriteLine($"Remaining Duplicate Order IDs: {remainingDuplicateIDs}");
                    }

                    using (XLWorkbook outputWorkbook = new XLWorkbook())
                    {
                        outputWorkbook.Worksheets.Add(cleanedTable, "Cleaned Data");
                        outputWorkbook.SaveAs(outputFile);
                    }

                    Console.WriteLine($"\nCleaned Dataset Saved As:");
                    Console.WriteLine(outputFile);

                    Console.WriteLine("\nPROJECT COMPLETED SUCCESSFULLY!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("\nERROR:");
                Console.WriteLine(ex.Message);
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
}