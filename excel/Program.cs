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
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var arr = new List<TimeValue>();
            string filePath = "/Users/magnusallerup/Documents/names.xlsx";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        var resTime = reader.GetValue(0);
                        var resValue = reader.GetValue(1);

                        if (resTime is double)
                        {
                            arr.Add(new TimeValue((double)resTime, (double)resValue));
                        }


                    }
                } while (reader.NextResult());
            }

            arr.ForEach(x => Console.WriteLine($"Time: {x.time}, Value: {x.value}"));

            var a = arr.Average(x => x.value);
            Console.WriteLine($"Average value: {a}");
            Console.Read();
        }

        public class TimeValue
        {
            public double time { get; set; }
            public double value { get; set; }

            public TimeValue(double x, double y)
            {
                time = x;
                value = y;
            }
        }
    }
}
