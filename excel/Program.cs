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
            //TimeValueArr();
            ISO17665();
        }

        public static void TimeValueArr()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var timeValueArr = new List<TimeValue>();
            var timeArr = new List<String>();
            var valueArr = new List<Double>();
            string filePath = "C:/Users/mra/Documents/wery7232XH.xlsx";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        var time = reader.GetValue(0);
                        var timeType = reader.GetFieldType(0);
                        var value = reader.GetValue(2);
                        var valueType = reader.GetFieldType(2);

                        if (time is string && value is double)
                        {
                            timeArr.Add((string)time);
                            timeValueArr.Add(new TimeValue((string)time, (double)value));
                        }
                    }
                } while (reader.NextResult());
            }
            //timeArr.ForEach(x => Console.WriteLine(x));
            //timeValueArr.ForEach(x => Console.WriteLine($"Time: {x.time}, Value: {x.value}"));

            Console.Read();
        }
        public static void ISO17665()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = "C:/Users/mra/Documents/test.xlsx";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                        object col1 = reader.GetValue(0);
                        var col1Type = reader.GetFieldType(0);
                        
                        object col3 = reader.GetValue(2);
                        var col3Type = reader.GetFieldType(0);
                        
                        object col5 = reader.GetValue(4);
                        var col5Type = reader.GetFieldType(0);

                        var row = reader.Depth;

                        PrintCol(row, col5Type, col5);
                    }
                } while (reader.NextResult());
            }
            Console.Read();
        }

        public static void PrintCol(int col, Type type, object value)
        {
            if (value != null)
            {
                Console.WriteLine($"Row {col + 1}: {type}: {value}");
            }
        }
        
        public class TimeValue
        {
            public string time { get; set; }
            public double value { get; set; }

            public TimeValue(string x, double y)
            {
                time = x;
                value = y;
            }
        }
    }
}
