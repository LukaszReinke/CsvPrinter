using System;
using System.Configuration;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;

namespace CsvPrinter
{
    class Program
    {
        private static string PrinterName;
        private static List<string> Headers;
        private static List<List<string>> Rows;
        private static int CurrentRowIndex;

        static void Main(string[] args)
        {
            Console.WriteLine("Podaj nazwę pliku");

            var filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' not found.");
                return;
            }

            PrinterName = ConfigurationManager.AppSettings["PrinterName"];
            if (string.IsNullOrWhiteSpace(PrinterName))
            {
                Console.WriteLine("Error: Printer name is not set in app.config.");
                return;
            }

            try
            {
                ReadCsv(filePath);
                CurrentRowIndex = 0; // Start from the first row
                PrintData();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void ReadCsv(string filePath)
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
            {
                csv.Read();
                csv.ReadHeader();
                Headers = csv.HeaderRecord.ToList();
                Rows = new List<List<string>>();

                while (csv.Read())
                {
                    var row = Headers.Select(header => csv.GetField(header)).ToList();
                    Rows.Add(row);
                }
            }
        }

        static void PrintData()
        {
            var printDoc = new PrintDocument();
            printDoc.PrinterSettings.PrinterName = PrinterName;
            printDoc.DefaultPageSettings.Landscape = true;
            printDoc.DefaultPageSettings.Margins = new Margins(10, 10, 10, 10); // Minimal margins

            if (!printDoc.PrinterSettings.IsValid)
            {
                Console.WriteLine($"Error: Printer '{PrinterName}' is not valid.");
                return;
            }

            printDoc.PrintPage += (sender, e) =>
            {
                var graphics = e.Graphics;
                float xMargin = e.MarginBounds.Left;
                float yMargin = e.MarginBounds.Top;
                float pageWidth = e.MarginBounds.Width * 0.98f; // Slight scale-down for width
                float pageHeight = e.MarginBounds.Height * 0.98f; // Slight scale-down for height
                float yPosition = yMargin;

                // Calculate column widths based on the longest text in each column
                var columnWidths = Headers.Select((header, colIndex) =>
                {
                    var maxContentLength = Rows.Max(row => row[colIndex].Length);
                    var maxLength = Math.Max(header.Length, maxContentLength);
                    return Math.Max(30, maxLength * 7); // Minimum width of 30 pixels
                }).ToList();

                // Adjust column widths based on the total page width
                float totalWidth = columnWidths.Sum();
                if (totalWidth > pageWidth)
                {
                    float scaleFactor = pageWidth / totalWidth;
                    columnWidths = columnWidths.Select(w => (int)(w * scaleFactor)).ToList();
                }

                // Adjust font size to a fixed small size for fitting more content
                var font = new Font("Arial", 6);
                float rowHeight = font.GetHeight() * 1.5f;

                // Draw table with borders
                float xPosition = xMargin;

                // Draw header with borders
                for (int i = 0; i < Headers.Count; i++)
                {
                    graphics.DrawRectangle(Pens.Black, xPosition, yPosition, columnWidths[i], rowHeight);
                    graphics.DrawString(Headers[i], font, Brushes.Black, new RectangleF(xPosition + 2, yPosition + 2, columnWidths[i] - 4, rowHeight - 4));
                    xPosition += columnWidths[i];
                }
                yPosition += rowHeight;

                // Draw rows with borders
                while (CurrentRowIndex < Rows.Count)
                {
                    var row = Rows[CurrentRowIndex];
                    xPosition = xMargin;

                    for (int i = 0; i < row.Count; i++)
                    {
                        graphics.DrawRectangle(Pens.Black, xPosition, yPosition, columnWidths[i], rowHeight);
                        graphics.DrawString(row[i], font, Brushes.Black, new RectangleF(xPosition + 2, yPosition + 2, columnWidths[i] - 4, rowHeight - 4));
                        xPosition += columnWidths[i];
                    }
                    yPosition += rowHeight;

                    if (yPosition + rowHeight > pageHeight)
                    {
                        e.HasMorePages = true;
                        CurrentRowIndex++;
                        return;
                    }

                    CurrentRowIndex++;
                }

                e.HasMorePages = false;
            };

            try
            {
                printDoc.Print();
                Console.WriteLine("Print job sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
