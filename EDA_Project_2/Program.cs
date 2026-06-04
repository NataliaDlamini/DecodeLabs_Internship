// ================================================
// DATA ANALYTICS PROJECT 2 - EDA IN C#
// ================================================

// Required NuGet Packages:
// Install-Package ExcelDataReader
// Install-Package ExcelDataReader.DataSet
// Install-Package ScottPlot

using System;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;
using ScottPlot;

class Program
{
    static void Main(string[] args)
    {
        // ============================================
        // LOAD EXCEL FILE
        // ============================================

        System.Text.Encoding.RegisterProvider(
            System.Text.CodePagesEncodingProvider.Instance);

        string filePath = "Dataset for Data Analytics (1).xlsx";

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

        using var reader = ExcelReaderFactory.CreateReader(stream);

        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true
            }
        });

        DataTable table = result.Tables[0];

        Console.WriteLine("================================");
        Console.WriteLine("DATA PREVIEW");
        Console.WriteLine("================================");

        // Display first 5 rows
        for (int i = 0; i < Math.Min(5, table.Rows.Count); i++)
        {
            foreach (DataColumn col in table.Columns)
            {
                Console.Write($"{table.Rows[i][col]} \t");
            }
            Console.WriteLine();
        }

        // ============================================
        // DESCRIPTIVE STATISTICS
        // ============================================

        Console.WriteLine("\n================================");
        Console.WriteLine("DESCRIPTIVE STATISTICS");
        Console.WriteLine("================================");

        foreach (DataColumn column in table.Columns)
        {
            // Check numeric columns only
            bool isNumeric = table.AsEnumerable()
                                  .All(r => double.TryParse(r[column].ToString(), out _));

            if (isNumeric)
            {
                var values = table.AsEnumerable()
                                  .Select(r => Convert.ToDouble(r[column]))
                                  .ToList();

                double mean = values.Average();
                double median = CalculateMedian(values);
                int count = values.Count;
                double min = values.Min();
                double max = values.Max();

                Console.WriteLine($"\nColumn: {column.ColumnName}");
                Console.WriteLine($"Count  : {count}");
                Console.WriteLine($"Mean   : {mean:F2}");
                Console.WriteLine($"Median : {median:F2}");
                Console.WriteLine($"Min    : {min}");
                Console.WriteLine($"Max    : {max}");

                // ============================================
                // OUTLIER DETECTION USING IQR
                // ============================================

                values.Sort();

                double q1 = Percentile(values, 25);
                double q3 = Percentile(values, 75);
                double iqr = q3 - q1;

                double lowerBound = q1 - 1.5 * iqr;
                double upperBound = q3 + 1.5 * iqr;

                var outliers = values
                    .Where(v => v < lowerBound || v > upperBound)
                    .ToList();

                Console.WriteLine($"Outliers Detected: {outliers.Count}");

                // ============================================
                // GENERATE HISTOGRAM
                // ============================================

                var plt = new ScottPlot.Plot();

                plt.Add.Histogram(values.ToArray());

                plt.Title($"Distribution of {column.ColumnName}");
                plt.XLabel(column.ColumnName);
                plt.YLabel("Frequency");

                string imageName = $"{column.ColumnName}_Histogram.png";

                plt.SavePng(imageName, 800, 600);

                Console.WriteLine($"Histogram Saved: {imageName}");
            }
        }

        Console.WriteLine("\nEDA COMPLETED SUCCESSFULLY!");
    }

    // ============================================
    // MEDIAN FUNCTION
    // ============================================

    static double CalculateMedian(List<double> numbers)
    {
        var sorted = numbers.OrderBy(n => n).ToList();

        int count = sorted.Count;

        if (count % 2 == 0)
        {
            return (sorted[count / 2 - 1] + sorted[count / 2]) / 2;
        }
        else
        {
            return sorted[count / 2];
        }
    }

    // ============================================
    // PERCENTILE FUNCTION
    // ============================================

    static double Percentile(List<double> sequence, double percentile)
    {
        double realIndex = percentile / 100.0 * (sequence.Count - 1);

        int index = (int)realIndex;

        double frac = realIndex - index;

        if (index + 1 < sequence.Count)
        {
            return sequence[index] * (1 - frac) + sequence[index + 1] * frac;
        }
        else
        {
            return sequence[index];
        }
    }
}
