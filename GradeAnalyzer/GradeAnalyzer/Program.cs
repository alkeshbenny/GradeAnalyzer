using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string basePath = @"C:\Users\alkes\OneDrive\Desktop\PROJECTS\GradeAnalyzer\GradeAnalyzer";

        string inputPath = Path.Combine(basePath, "Grades.xlsx");
        string outputPath = Path.Combine(basePath, "Results_ClosedXML_Only.xlsx");

        if (!File.Exists(inputPath))
        {
            Console.WriteLine("Input Excel file not found.");
            return;
        }

        Console.WriteLine("Reading and processing Excel file using ClosedXML...");
        var studentResults = ReadAndCalculateWithClosedXML(inputPath);

        Console.WriteLine("Writing average grades using ClosedXML...");
        WriteResultsWithClosedXML(outputPath, studentResults);

        Console.WriteLine("Processing complete.");
    }

    public class StudentResult
    {
        public string Name { get; set; }
        public double Average { get; set; }
        public string Status => Average >= 60 ? "Pass" : "Fail";
    }

    // 📥 Read + calculate using ClosedXML
    static List<StudentResult> ReadAndCalculateWithClosedXML(string filePath)
    {
        var results = new List<StudentResult>();

        using (var workbook = new XLWorkbook(filePath))
        {
            var sheet = workbook.Worksheet(1); // First sheet
            var rows = sheet.RowsUsed().Skip(1); // Skip header

            foreach (var row in rows)
            {
                string name = row.Cell(1).GetValue<string>();
                var grades = row.Cells(2, row.LastCellUsed().Address.ColumnNumber)
                                .Select(cell => cell.GetDouble())
                                .ToList();

                double average = grades.Any() ? grades.Average() : 0;
                results.Add(new StudentResult { Name = name, Average = average });
            }
        }

        return results;
    }

    // 📤 Write results using ClosedXML
    static void WriteResultsWithClosedXML(string filePath, List<StudentResult> results)
    {
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Results");
            sheet.Cell(1, 1).Value = "Student Name";
            sheet.Cell(1, 2).Value = "Average";
            sheet.Cell(1, 3).Value = "Status";

            int row = 2;
            foreach (var student in results)
            {
                sheet.Cell(row, 1).Value = student.Name;
                sheet.Cell(row, 2).Value = Math.Round(student.Average, 2);
                sheet.Cell(row, 3).Value = student.Status;



                row++;
            }

            workbook.SaveAs(filePath);
        }
    }
}
