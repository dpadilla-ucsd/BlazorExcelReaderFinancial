using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorExcelReaderFinancial.Data
{
    public class FinancialData
    {
        public string CellData1 { get; set; } = "";
        public string CellData2 { get; set; } = "";
        public string CellData3 { get; set; } = "";
        public string CellData4 { get; set; } = "";
        public string CellData5 { get; set; } = "";
        public string CellData6 { get; set; } = "";


        public List<FinancialData> ReadExcel()
        {
            List<FinancialData> finances = new List<FinancialData>();

            string FilePath = @"C:\Users\davidp\Documents\FINAL-Fiscal Workbook-Year1-test.xlsx";
            FileInfo existingFile = new FileInfo(FilePath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                System.Diagnostics.Debug.Print("Current worksheet:" + worksheet.Name);
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;
                //FinancialData dataPoint = new FinancialData();
                System.Diagnostics.Debug.Print("Cols:" + colCount + "  Rows:" + rowCount);
                for (int row = 1; row <= rowCount; row++)
                {
                    FinancialData dataPoint = new FinancialData();
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue;
                        if (worksheet.Cells[row, col].Text.ToString().Length > 0)
                        {
                            cellValue = worksheet.Cells[row, col].Text;
                        }
                        else
                        {
                            cellValue = "";
                        }

                        switch (col)
                        {
                            case 1:
                                dataPoint.CellData1 = cellValue;
                                break;
                            case 2:
                                dataPoint.CellData2 = cellValue;
                                break;
                            case 3:
                                dataPoint.CellData3 = cellValue;
                                break;
                            case 4:
                                dataPoint.CellData4 = cellValue;
                                break;
                            case 5:
                                dataPoint.CellData5 = cellValue;
                                break;
                            case 6:
                                dataPoint.CellData6 = cellValue;
                                break;
                            default:
                                System.Diagnostics.Debug.Print("Default case");
                                break;
                        }

                        System.Diagnostics.Debug.Print("Cell data read:" + cellValue);
                        //finances.Add(dataPoint);
                    } // end for rows
                    finances.Add(dataPoint);
                } // end for col
            }
            return finances;
        }
    }
}
