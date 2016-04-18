using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;

namespace EnergyPlusConverter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("This program takes a csv output from EnergyPlus and turns it into a converted .xlsx file for further analysis.");

            FileDialog fileDialog = new OpenFileDialog();

            DialogResult result = fileDialog.ShowDialog();

            if (result != DialogResult.OK) 
            {
                Console.WriteLine("No file selected, press any key to exit...");
                Console.ReadLine();
                return;
            }

            string file = fileDialog.FileName;

            if (string.IsNullOrEmpty(file))
            {
                Console.WriteLine("No file was selected. Exiting.");
                Console.ReadLine();
                return;
            }

            FileInfo fileInfo = new FileInfo(file);
            
            ExcelPackage excelFile = new ExcelPackage();

            ExcelWorksheet excelWorksheet = excelFile.Workbook.Worksheets.Add("Converted Data");

            try
            {
                excelWorksheet.Cells[1, 1].LoadFromText(fileInfo);
            }
            catch (Exception)
            {
                Console.WriteLine("There was an issue loading the .csv file. Are you this is a csv file from EnergyPlus?");
                Console.WriteLine("Press any key to exit.");
                Console.ReadLine();
                return;
            }
            excelWorksheet.Row(1).Style.WrapText = true;

            excelWorksheet.Column(1).Style.Numberformat.Format = "mm/dd hh:MM";

            int totalRows = excelWorksheet.Dimension.Rows;
            int totalCols = excelWorksheet.Dimension.Columns;
            for (int col = 1; col <= excelWorksheet.Dimension.Columns; col++)
            {
                ExcelRange headerCell = excelWorksheet.Cells[1, col];
                string headerString = headerCell.Value.ToString();
                if (headerString.Contains("Temperature") && headerString.Contains("[C]"))
                {
                    ExcelRange cells = excelWorksheet.Cells[2, col, totalRows, col];
                    foreach (ExcelRangeBase cell in cells)
                    {
                        cell.Value = double.Parse(cell.Value.ToString())*9/5 + 32;
                    }

                    headerCell.Value = headerString.Replace("[C]", "[F]");
                    excelWorksheet.Column(col).Style.Numberformat.Format = "0.0";
                }

                if (headerString.Contains("Electricity") && headerString.Contains("[J]"))
                {
                    ExcelRange cells = excelWorksheet.Cells[2, col, totalRows, col];
                    foreach (ExcelRangeBase cell in cells)
                    {
                        cell.Value = double.Parse(cell.Value.ToString())/3600000.0;
                    }
                    headerCell.Value = headerString.Replace("[J]", "[kWh]");

                    excelWorksheet.Column(col).Style.Numberformat.Format = "#,##0";
                }
                if ((headerString.Contains("DistrictHeating") || headerString.Contains("DistrictCooling")) && headerString.Contains("[J]"))
                {
                    ExcelRange cells = excelWorksheet.Cells[2, col, totalRows, col];
                    foreach (ExcelRangeBase cell in cells)
                    {
                        cell.Value = double.Parse(cell.Value.ToString())/1055.06/1000000.0;
                    }
                    headerCell.Value = headerString.Replace("[J]", "[MMBTU]");

                    excelWorksheet.Column(col).Style.Numberformat.Format = "#,##0";
                }


            }

            for (int i = 1; i <= totalCols; i++)
            {
                excelWorksheet.Column(i).Width = 15;
            }

            string baseFileName = fileInfo.Directory + "\\" + Path.GetFileNameWithoutExtension(file) + "-convert.xlsx";
            string trialFileName = baseFileName;

            int num = 1;
            while (File.Exists(trialFileName))
            {
                trialFileName = fileInfo.Directory + "\\" + Path.GetFileNameWithoutExtension(baseFileName) + num + ".xlsx";
                num++;
            }

            FileInfo newFile = new FileInfo(trialFileName);

            excelFile.SaveAs(newFile);

            Console.WriteLine("Successfully converted {0} to {1}.", fileInfo.Name, newFile.Name);
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
            return;
        }
    }
}
