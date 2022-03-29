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

        if (Path.GetExtension(filePath).ToUpper() == ".XLS")
        {
            excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
        }
        else
        {      
            excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }

        DataSet result = excelReader.AsDataSet();

        //excelReader.IsFirstRowAsColumnNames = false;
        DataTable dt = result.Tables[0];
        Console.WriteLine(result.Tables[0].Rows.Count);
        Console.WriteLine(result.Tables[0].Columns.Count);
        excelReader.Close();
    }
}
