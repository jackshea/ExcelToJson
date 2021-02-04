using System;
using System.Text;
using ExcelToJsonLib;
using Newtonsoft.Json;

namespace ExcelToJsonConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var excelToJson = new ExcelToJson();
            string json = excelToJson.OpenExcelAndToJson("../../../Res/Examples/ExampleData.xlsx");
            Console.WriteLine(json);
        }
    }
}
