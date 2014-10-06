using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.Data;

namespace ExcelToCSVDemo
{
    public class ExcelToCSV
    {
        
        private enum FileExtensions { xls, xlsx };
        private FileExtensions excelExtension;
        private string excelFilename;

        public ExcelToCSV(string pathToExcelFile)
        {
            setFileProperties(pathToExcelFile);
        }
        
        private void setFileProperties(string filepath)
        {
            bool extensionFound = false;

            if (String.IsNullOrEmpty(filepath) || String.IsNullOrWhiteSpace(filepath))
            {
                throw new Exception("The given filepath is null or empty");
            }
            else if (filepath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                excelExtension = FileExtensions.xls;
                extensionFound = true;
            }
            else if (!filepath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                excelExtension = FileExtensions.xlsx;
                extensionFound = true;
            }
            else if (!extensionFound)
            {
                throw new Exception("The file extension is not expected.");
            }
            else if (!File.Exists(excelFilename))
            {
                throw new FileNotFoundException();
            }
            else
            {
                excelFilename = filepath;
            }
        }

        public DataSet readExcelFile()
        {
            DataSet dataset = new DataSet();
            FileStream fileStream = new FileStream(excelFilename, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = null;

            try
            {
                if (excelExtension == FileExtensions.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                else if (excelExtension == FileExtensions.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

                excelReader.IsFirstRowAsColumnNames = true;
                dataset = excelReader.AsDataSet();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                fileStream.Close();

                if (excelReader != null)
                    excelReader.Close();
            }

            return dataset;
        }

        public void writeDataToCSVFile(DataSet excelWorkbook)
        {
            if (excelWorkbook == null)
            {
                throw new ArgumentNullException();
            }
            else if (excelWorkbook.Tables.Count == 0)
            {
                throw new Exception("The Excel Workbook has no spreadsheets or tables.");
            }
            else
            {
                foreach (DataTable dt in excelWorkbook.Tables)
                {
                    string spreadsheetName = dt.TableName;

                }
            }
        }
    }
}
