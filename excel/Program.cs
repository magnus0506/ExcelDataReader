using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using System.Linq;
namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            ISO17665();
        }


        //Testprotokol ISO17665
        public static void ISO17665()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = "excelArk.xlsx";
            var col1Arr = new List<double>();
            var col2Arr = new List<double>();
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        col1Arr.Add((double)reader.GetValue(0));
                        col2Arr.Add((double)reader.GetValue(1));
                    }
                } while (reader.NextResult());
            }
            Console.WriteLine($"Første kolonne - Observationer: {col1Arr.Count()}   Gennemsnit: {col1Arr.Average()}" +
                              $"   Min. værdi: {col1Arr.Min()}  Max. værdi: {col1Arr.Max()}\n" +
                              $"Anden kolonne - Observationer: {col2Arr.Count()}   Gennemsnit: {col2Arr.Average()}" +
                              $"   Min. værdi: {col2Arr.Min()}  Max. værdi: {col2Arr.Max()}");
            Console.Read();
        }
    }
}
