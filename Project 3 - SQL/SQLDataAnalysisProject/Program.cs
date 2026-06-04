using System;
using System.Data;
using System.IO;
using Microsoft.Data.Sqlite;
using ExcelDataReader;

namespace DataAnalyticsProject3
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(
                System.Text.CodePagesEncodingProvider.Instance);

            string excelFile = "Dataset for Data Analytics.xlsx";
            string databaseFile = "SalesAnalytics.db";

            using var connection =
                new SqliteConnection($"Data Source={databaseFile}");

            connection.Open();

            CreateTable(connection);
            ImportExcelData(connection, excelFile);

            Console.WriteLine("=================================");
            Console.WriteLine("SQL DATA ANALYSIS REPORT");
            Console.WriteLine("=================================");

            ExecuteQuery(connection,
                "SELECT COUNT(*) AS TotalOrders FROM Orders;",
                "TOTAL ORDERS");

            ExecuteQuery(connection,
                @"SELECT Product,
                         SUM(Quantity) AS TotalQuantitySold
                  FROM Orders
                  GROUP BY Product
                  ORDER BY TotalQuantitySold DESC;",
                "TOP SELLING PRODUCTS");

            ExecuteQuery(connection,
                @"SELECT PaymentMethod,
                         COUNT(*) AS NumberOfOrders
                  FROM Orders
                  GROUP BY PaymentMethod;",
                "ORDERS BY PAYMENT METHOD");

            ExecuteQuery(connection,
                @"SELECT AVG(TotalPrice) AS AverageOrderValue
                  FROM Orders;",
                "AVERAGE ORDER VALUE");

            ExecuteQuery(connection,
                @"SELECT OrderStatus,
                         COUNT(*) AS Total
                  FROM Orders
                  GROUP BY OrderStatus;",
                "ORDER STATUS DISTRIBUTION");

            ExecuteQuery(connection,
                @"SELECT ReferralSource,
                         SUM(TotalPrice) AS Revenue
                  FROM Orders
                  GROUP BY ReferralSource
                  ORDER BY Revenue DESC;",
                "REVENUE BY REFERRAL SOURCE");

            Console.WriteLine("\nAnalysis Complete.");
        }

        static void CreateTable(SqliteConnection connection)
        {
            string sql = @"
            CREATE TABLE IF NOT EXISTS Orders
            (
                OrderID TEXT,
                Date TEXT,
                CustomerID TEXT,
                Product TEXT,
                Quantity INTEGER,
                UnitPrice REAL,
                ShippingAddress TEXT,
                PaymentMethod TEXT,
                OrderStatus TEXT,
                TrackingNumber TEXT,
                ItemsInCart INTEGER,
                CouponCode TEXT,
                ReferralSource TEXT,
                TotalPrice REAL
            );";

            new SqliteCommand(sql, connection).ExecuteNonQuery();
        }

        static void ImportExcelData(
            SqliteConnection connection,
            string excelFile)
        {
            using var stream =
                File.Open(excelFile, FileMode.Open, FileAccess.Read);

            using var reader =
                ExcelReaderFactory.CreateReader(stream);

            var result = reader.AsDataSet();

            DataTable table = result.Tables[0];

            for (int i = 1; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];

                string sql = @"
                INSERT INTO Orders
                VALUES
                (
                    @OrderID,
                    @Date,
                    @CustomerID,
                    @Product,
                    @Quantity,
                    @UnitPrice,
                    @ShippingAddress,
                    @PaymentMethod,
                    @OrderStatus,
                    @TrackingNumber,
                    @ItemsInCart,
                    @CouponCode,
                    @ReferralSource,
                    @TotalPrice
                );";

                using var cmd =
                    new SqliteCommand(sql, connection);

                cmd.Parameters.AddWithValue("@OrderID", row[0]);
                cmd.Parameters.AddWithValue("@Date", row[1]);
                cmd.Parameters.AddWithValue("@CustomerID", row[2]);
                cmd.Parameters.AddWithValue("@Product", row[3]);
                cmd.Parameters.AddWithValue("@Quantity", row[4]);
                cmd.Parameters.AddWithValue("@UnitPrice", row[5]);
                cmd.Parameters.AddWithValue("@ShippingAddress", row[6]);
                cmd.Parameters.AddWithValue("@PaymentMethod", row[7]);
                cmd.Parameters.AddWithValue("@OrderStatus", row[8]);
                cmd.Parameters.AddWithValue("@TrackingNumber", row[9]);
                cmd.Parameters.AddWithValue("@ItemsInCart", row[10]);
                cmd.Parameters.AddWithValue("@CouponCode", row[11]);
                cmd.Parameters.AddWithValue("@ReferralSource", row[12]);
                cmd.Parameters.AddWithValue("@TotalPrice", row[13]);

                cmd.ExecuteNonQuery();
            }

            Console.WriteLine("Excel data imported successfully.");
        }

        static void ExecuteQuery(
            SqliteConnection connection,
            string query,
            string title)
        {
            Console.WriteLine($"\n===== {title} =====");

            using var cmd =
                new SqliteCommand(query, connection);

            using var reader =
                cmd.ExecuteReader();

            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    Console.Write(
                        $"{reader.GetName(i)}: {reader.GetValue(i)}   ");
                }

                Console.WriteLine();
            }
        }
    }
}