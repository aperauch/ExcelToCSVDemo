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
            Console.WriteLine("Enter the complete filepath to the Excel file to be converted:");

            bool run = true;
            while (run)
            {
                string input = Console.ReadLine(); 
                
                if (input.Equals("q", StringComparison.OrdinalIgnoreCase))
                {
                    run = false;
                    Console.WriteLine("Now quiting.");
                }
                else
                {
                   ExcelToCSV converter = new ExcelToCSV(input);

                   if (converter.readyForRead)
                   {
                       DataSet excelWorkbook = converter.readExcelFile();
                       converter.writeDataToCSVFile(excelWorkbook);
                   }
                }

                Console.WriteLine("Enter an Excel filepath or press 'Q' to quit:");
            }
        }
    }
}
