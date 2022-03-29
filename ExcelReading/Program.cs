using System;
using ExcelDataReader;
using System.Data;

class Program
{
    public static void Main(string[] args)
    {
        string filePath = @"C:\Users\Izagakhmaevra\Desktop\Excel\TestExel.xlsx";

        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        IExcelDataReader excelReader;

        // Чтение Excel файла
        if (Path.GetExtension(filePath).ToUpper() == ".XLS")
        {
            // Если фомат файла = .XLS
            excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
        }
        else
        {
            // Или формат файла = .XLSX
            excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        DataSet result = excelReader.AsDataSet();

        excelReader.IsFirstRowAsColumnNames = false;

        DataTable dt = result.Tables[0];
        Console.WriteLine(result.Tables[0].Rows.Count);
        Console.WriteLine(result.Tables[0].Columns.Count);

        excelReader.Close();
        //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        //{
        //    IExcelDataReader reader;
        //    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

        //    var conf = new ExcelDataSetConfiguration
        //    {
        //        ConfigureDataTable = _ => new ExcelDataTableConfiguration
        //        {
        //            UseHeaderRow = true
        //        }
        //    };
        //    var dataSet = reader.AsDataSet(conf);

        //    // Now you can get data from each sheet by its index or its "name"
        //    var dataTable = dataSet.Tables[0];
        //    Console.WriteLine(dataSet);
        //}
    }
}
