using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;

namespace coreconsole
{
    class Program
    {
        static List<JumboRate> jumboRates = new List<JumboRate>();

        static void Main(string[] args)
        {
            var sourceUrl = "https://www.naic.org/documents/pbr_data_2018_vm-22_non-jumbo_and_jumbo_valuation_rates.xlsx";
            var tempFile = Path.GetTempFileName();
            var webClient = new WebClient();

            Console.WriteLine("Downloading...");
            webClient.DownloadFile(sourceUrl, tempFile);

            Console.WriteLine("Parsing...");
            using (var package = new ExcelPackage(new FileInfo(tempFile)))
            {
                var sheet1 = package.Workbook.Worksheets[0];

                for (var row = 5; cellHasValue(sheet1.Cells[row,2].Value); row++)
                {
                    var date = (DateTime)sheet1.Cells[row, 1].Value;
                    Console.WriteLine($"Processing {date:d}");

                    var bucketA = (decimal)(double)sheet1.Cells[row, 2].Value;
                    var bucketB = (decimal)(double)sheet1.Cells[row, 3].Value;
                    var bucketC = (decimal)(double)sheet1.Cells[row, 4].Value;
                    var bucketD = (decimal)(double)sheet1.Cells[row, 5].Value;

                    jumboRates.Add(new JumboRate(date, bucketA, bucketB, bucketC, bucketD));
                }
            }

            Console.WriteLine($"Avg rate for BucketA is {jumboRates.Select(r => r.BucketA).Average()}");

            Console.WriteLine("Cleaning up.");
            File.Delete(tempFile);
        }

        static bool cellHasValue(object cellValue)
        {
            return (cellValue?.ToString() ?? String.Empty).Trim() != String.Empty;
        }
    }

    public class JumboRate
    {
        public JumboRate(DateTime date, decimal bucketA, decimal bucketB, decimal bucketC, decimal bucketD)
        {
            Date = date;
            BucketA = bucketA;
            BucketB = bucketB;
            BucketC = bucketC;
            BucketD = bucketD;
        }

        public DateTime Date { get; }
        public decimal BucketA { get; }
        public decimal BucketB { get; }
        public decimal BucketC { get; }
        public decimal BucketD { get; }
    }
}
