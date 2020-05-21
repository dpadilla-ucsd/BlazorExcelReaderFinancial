using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorExcelReaderFinancial.Data
{
    public class FileTab
    {
        public string tabName = "";

        public int tabNumber;

        public List<FileTab> GetFileTabs()
        {
            List<FileTab> fileTabs = new List<FileTab>();

            string FilePath = @"C:\Users\davidp\Documents\FINAL-Fiscal Workbook-Year1-test.xlsx";
            FileInfo existingFile = new FileInfo(FilePath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                int tabCount = package.Workbook.Worksheets.Count();
                System.Diagnostics.Debug.Print("~~~~ READING " + tabCount + " FILE TABS ~~~");
                for (int tab = 0; tab < tabCount; tab++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[tab];
                    System.Diagnostics.Debug.Print(tab + ":" + worksheet.Name);
                    FileTab fileTab = new FileTab();
                    fileTab.tabName = worksheet.Name;
                    fileTab.tabNumber = tab;
                    fileTabs.Add(fileTab);
                }
            }
            return fileTabs;
        }
    }
}
