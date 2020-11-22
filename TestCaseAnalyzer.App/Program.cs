using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace TestCaseAnalyzer.App
{
    class Program
    {
        static void Main(string[] args)
        {
            Directory.CreateDirectory("report");
            Directory.CreateDirectory("ExcelReport");
            Directory.CreateDirectory("ExcelReportwithAWorkBook");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new DataFileReader();
            var testCases = reader.ReadTestCases().ToList();

            var requirements = reader.ReadRequirements().ToList();

            string menu = null;

            for (int i = 0; i < testCases.Count; i++)
            {
                TestCase testCase = testCases[i];
                if (testCase.ID != null)
                {
                    menu += $"<a href='{testCase.ID.Replace("#", "")}.html'>{testCase.ID}</a><br>\n";
                }
            }

            foreach (var testCase in testCases)
            {
                if (testCase.ID != null)
                {
                    string testCaseDetail = TestCaseHtmlGenerator.CreateTestCaseHtml(
                        requirements,
                        testCase);

                    File.WriteAllText(
                        $"report/{testCase.ID.Replace("#", "")}.html",
                        testCaseDetail);
                }
            }

            Console.WriteLine(menu);

            File.WriteAllText("report/index.html", menu);


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

            //Create new Excel WorkBook document. 
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
