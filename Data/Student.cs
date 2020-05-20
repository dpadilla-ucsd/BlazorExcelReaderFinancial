using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorExcelReaderFinancial.Data
{
    public class Student
    {
        public string FirstName { get; set; } = "";

        public string LastName { get; set; } = "";

        public string FullName => this.FirstName + " " + this.LastName;

        public double StudentNumber { get; set; } = 0;

        public List<Student> ReadExcel()
        {
            List<Student> students = new List<Student>();

            string FilePath = @"C:\Users\davidp\Documents\Document.xlsx";
            FileInfo existingFile = new FileInfo(FilePath);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

                int tabCount = package.Workbook.Worksheets.Count();
                System.Diagnostics.Debug.Print("Number of tabs=" + tabCount);
                for (int tab = 0; tab < tabCount; tab++)
                {
                    System.Diagnostics.Debug.Print(tab + ":" + package.Workbook.Worksheets[tab].Name);
                }

                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 1; row <= rowCount; row++)
                {
                    Student student = new Student();
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (col == 1) student.FirstName = worksheet.Cells[row, col].Value.ToString();
                        else if (col == 2) student.LastName = worksheet.Cells[row, col].Value.ToString();
                        else if (col == 3) student.StudentNumber = (double)(worksheet.Cells[row, col].Value);
                    }
                    students.Add(student);
                }
            }

            return students;
        }
    }
}
