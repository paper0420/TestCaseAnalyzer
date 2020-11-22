using IronXL;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class ExcelReportWithAWorkbookGenerator : IReportGenerator
    {
        public void GenerateReport(List<TestCase> testCases, List<Requirement> requirements)
        {
            Directory.CreateDirectory("ExcelReportwithAWorkBook");
            WorkBook xlsxWorkbook2 = WorkBook.Create(ExcelFileFormat.XLSX);

            List<string> testCaseID = new List<string>();

            foreach (var testCase in testCases)
            {

                if (testCase.ID != null)
                {
                    if (!testCaseID.Contains(testCase.ID))
                    {
                        WorkSheet xlsSheet2 = xlsxWorkbook2.CreateWorkSheet($"{testCase.ID}");
                        Console.WriteLine(xlsSheet2.Name);

                        TestCaseExcelGenerator.CreateTestCaseExcel(requirements, testCase, xlsSheet2);
                    }

                }

                testCaseID.Add(testCase.ID);

            }


            xlsxWorkbook2.SaveAs("ExcelReportwithAWorkBook/TestCaseAnalysis.xlsx");

        }
    }
}
