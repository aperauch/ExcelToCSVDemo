using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel;
using System.Text.RegularExpressions;

namespace ExcelToCSVDemo
{
    /// <summary>
    /// Converts spreadsheets contained in an Excel file to a CSV formatted file.  Uses the ExcelDataReader NuGet package.
    /// 
    /// ExcelDataReader:  https://github.com/ExcelDataReader
    /// </summary>
    public class ExcelToCSV
    {
        //Class properties
        public bool readyForRead { get; private set; }

        //Class variables
        private enum FileExtensions { xls, xlsx };
        private FileExtensions enumFileExtension;
        private string excelFilepath;
        private string excelDirectory;
        private string excelFilename;
        private string excelExtension;
        private string workbookName;
        private List<string> spreadsheetNames;

        /// <summary>
        /// Constructor for class.
        /// </summary>
        /// <param name="pathToExcelFile">The complete filepath to the Excel file.</param>
        public ExcelToCSV(string pathToExcelFile)
        {
            readyForRead = false;
            setFileProperties(pathToExcelFile);
        }
        
        /// <summary>
        /// Method sets the file properties of the specified file.
        /// </summary>
        /// <param name="filepath">The filepath of the file to be converted.</param>
        private void setFileProperties(string filepath)
        {
            if (String.IsNullOrEmpty(filepath) || String.IsNullOrWhiteSpace(filepath))
            {
                Console.Write("The given filepath is empty.  ");
            }
            else if (!File.Exists(filepath))
            {
                Console.Write("The file was not found.  ");
            }
            else
            {
                try
                {
                    excelExtension = Path.GetExtension(filepath);

                    if (excelExtension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                        enumFileExtension = FileExtensions.xls;
                    else if (excelExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                        enumFileExtension = FileExtensions.xlsx;
                    else
                        Console.WriteLine("The file extension {0} is not expected.  Will attempt to export.", excelExtension);

                    excelFilepath = filepath;
                    excelDirectory = Path.GetDirectoryName(filepath);
                    excelFilename = Path.GetFileNameWithoutExtension(filepath);
                    readyForRead = true;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        /// <summary>
        /// Method reads the contents of the Excel file into a DataSet object and stores in memory for processing.
        /// </summary>
        /// <returns>A DataSet containing all the data within the given Excel file.</returns>
        public DataSet readExcelFile()
        {
            DataSet dataset = new DataSet();
            FileStream fileStream = null;
            IExcelDataReader excelReader = null;

            try
            {
                fileStream = new FileStream(excelFilepath, FileMode.Open, FileAccess.Read);

                if (enumFileExtension == FileExtensions.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(fileStream);
                else if (enumFileExtension == FileExtensions.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

                workbookName = excelReader.Name;
                dataset = excelReader.AsDataSet();
                spreadsheetNames = getSpreadsheetNames(dataset);
            }
            catch (IOException ioe)
            {
                Console.WriteLine(ioe.Message);
                Console.WriteLine("Please close the file before proceeding:  " + excelFilepath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                if (excelReader != null)
                {
                    fileStream.Close();
                    excelReader.Close();
                }
            }

            return dataset;
        }

        /// <summary>
        /// Method returns the names of the Excel spreadsheets that are contained within the specified Excel workbook.
        /// </summary>
        /// <param name="excelWorkbook">The DataSet extracted from the Excel file.</param>
        /// <returns>A string list of the Excel spreadsheet names.</returns>
        private List<string> getSpreadsheetNames(DataSet excelWorkbook)
        {
            List<string> tableName = new List<string>();
            
            foreach (DataTable dt in excelWorkbook.Tables)
            {
                string formattedStr = dt.TableName.Replace(" ", "-");
                tableName.Add(formattedStr);
            }

            if (excelWorkbook.Tables.Count != tableName.Count)
                Console.WriteLine("Spreadsheet name and count mismatch.");

            return tableName;
        }

        /// <summary>
        /// Method returns a string list of the complete filepaths of the CSV files to be created.
        /// </summary>
        /// <param name="directory">The file directory to write the CSV files.</param>
        /// <param name="filename">The filename of the Excel workbook.</param>
        /// <param name="spreadsheetNames">The list of Excel spreadsheet names.</param>
        /// <returns>A string list of the complete filepath for each spreadsheet.</returns>
        private List<string> getCSVFilenames(string directory, string filename, List<string> spreadsheetNames)
        {
            List<string> csvFilenames = new List<string>();

            foreach (string sheetName in spreadsheetNames)
            {
                csvFilenames.Add(directory + "\\" + filename + "_" + sheetName + ".csv");
            }

            return csvFilenames;
        }

        /// <summary>
        /// Method writes the data stored in the specified DataSet to a CSV formatted files.
        /// </summary>
        /// <param name="excelWorkbook">The DataSet to be written to a file in CSV format.</param>
        public void writeDataToCSVFile(DataSet excelWorkbook)
        {
            if (excelWorkbook == null)
            {
                Console.WriteLine("The Excel workbook is null.");
            }
            else if (excelWorkbook.Tables.Count == 0)
            {
                Console.WriteLine("The Excel Workbook has no spreadsheets or tables.");
            }
            else
            {
                List<string> csvFilenames = getCSVFilenames(excelDirectory, excelFilename, spreadsheetNames);

                for (int i = 0; i < excelWorkbook.Tables.Count; i++)
                {
                    StringBuilder sb = new StringBuilder();

                    foreach (DataRow row in excelWorkbook.Tables[i].Rows)
                    {
                        sb.AppendLine(string.Join(",", row.ItemArray.Select(r => r.ToString()).ToArray()));
                    }

                    File.WriteAllText(csvFilenames[i], sb.ToString());
                }

                foreach (string name in csvFilenames)
                {
                    Console.WriteLine("Created file:  " + name);
                }
            }
        }
    }
}
