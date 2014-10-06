using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCSVDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the complete filepath to the Excel file to be converted:  ");
            string input = "c:\\Temp\\TestData.xlsx";//Console.ReadLine();

            ExcelToCSV converter = new ExcelToCSV(input);
            DataSet excelWorkbook = converter.readExcelFile();
            converter.writeDataToCSVFile(excelWorkbook);
        }
    }
}
