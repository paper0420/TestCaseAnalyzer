using IronXL;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class ExcelsReportGenerator : IReportGenerator
    {
        public void GenerateReport(List<TestCase> testCases, List<Requirement> requirements)
        {
            Directory.CreateDirectory("ExcelReport");
            foreach (var testCase in testCases)
            {

                if (testCase.ID != null)
                {

                    //Create new Excel WorkBook document. 
                    WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
                    WorkSheet xlsSheet = xlsxWorkbook.CreateWorkSheet($"{testCase.ID}");
                    TestCaseExcelGenerator.CreateTestCaseExcel(requirements, testCase, xlsSheet);
                    xlsxWorkbook.SaveAs($"ExcelReport/{testCase.ID.Replace("#", "")}.xlsx");

                }

            }
        }
    }
}
