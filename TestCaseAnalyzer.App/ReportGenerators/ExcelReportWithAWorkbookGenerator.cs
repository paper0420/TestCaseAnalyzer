using IronXL;
using System.Collections.Generic;
using System.IO;
using TestCaseAnalyzer.App.Spec;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public static class ExcelReportWithAWorkbookGenerator
    {
        public static void GenerateTestCaseAndRequirment(SpecParameters spec)
        {
            Directory.CreateDirectory("ExcelReportwithAWorkBook");
            WorkBook xlsxWorkbook2 = WorkBook.Create(ExcelFileFormat.XLSX);


            List<string> testCaseID = new List<string>();

            foreach (var testCase in spec.TestCases)
            {

                if (testCase.ID != null)
                {
                    if (!testCaseID.Contains(testCase.ID))
                    {
                        WorkSheet xlsSheet2 = xlsxWorkbook2.CreateWorkSheet($"{testCase.ID}");

                        //Console.WriteLine(xlsSheet2.Name);

                        TestcasesAndRequirementExcelGenerator.CreateTestCaseExcel(spec.CurrentRequirements, testCase, xlsSheet2);

                    }

                }

                testCaseID.Add(testCase.ID);

            }


            xlsxWorkbook2.SaveAs("ExcelReportwithAWorkBook/TestCaseAndRequirement.xlsx");

        }


        //public static void GenerateTestSpec(SpecParameters spec)
        //{
        //    Directory.CreateDirectory("TestSpec");
        //    WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
        //    spec.XlsSheet = xlsxWorkbook.CreateWorkSheet("Test_Item");

        //    Console.WriteLine("Gen folder");

        //    TestSpecExcelGenerator.CreateTestSpecExcel(spec);

        //    xlsxWorkbook.SaveAs("TestSpec/TestSpecification.xlsx");
        //}



    }
}
