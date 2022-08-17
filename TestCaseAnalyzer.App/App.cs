using System;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.FileReader;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    public class App
    {
        public static void RunMyApp(string newFile, string currentFile)
        {
            


            Console.Write("Enter Car Line: ");
            string carLine = Console.ReadLine();
            Console.Write("Enter report type: ");
            string reportType = Console.ReadLine();
           
            Console.WriteLine($"***Car Line: {carLine} Report type: {reportType}");


            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new ExcelReader();
            var currentKLH = reader.ReadFile(FileNames.KlhFile, "KLH_BL11.1", 1, (t,y) => new Requirement(t,y)).ToList();
            var executedTestcases = reader.ReadFile(FileNames.TestSpecFile, "Test_Item", 1, (t,y) => new TestCaseOnlyExecutedItem(t,y)).ToList();
            var htmlReports = HtmlReader.ReadHtmlFullReport(reportType).ToList();

            //var panaTestCases = reader.ReadFile("data1.xlsm", "5.テスト項目Test item", 18, t => new TestCase(t)).ToList();
            //var newKLH = reader.ReadFile(newFile, "Sheet1", 2, t => new Requirement(t)).ToList();
            //var delReqs = reader.ReadFile(newFile, "Sheet1", 2, t => new DeletedReq(t))
            //    .Where(t => t.ID != "0")
            //    .ToList();
            //var rejReqs = reader.ReadFile(newFile, "Sheet1", 2, t => new RejectedReq(t))
            //   .Where(t => t.ID != "0")
            //   .ToList();

            //var syrItems = reader.ReadFile("Copy of Eng02_to_Eng01_No_verification_criteria.xlsx", "Sheet1", 2, t => new SYRitem(t)).ToList();
            //var syrLists = reader.ReadFile("Copy of Eng02_to_Eng01_No_verification_criteria.xlsx", "TestCaseID_Linked(V2)", 2, t => new Dummy(t)).ToList();

            //var eng9TestCases = reader.ReadFile("ENG9_FuncTCs.xlsx", "ENG9_FuncTCs", 7, t => new ENG9_Func_TestCase(t)).ToList();
            //var klhLists = reader.ReadFile("Test_caseIDs.xlsx", "Sheet1", 2, t => new Dummy(t)).ToList();

            var baseSpec = new SpecParameters(
                currentRequirments: currentKLH,
                testCases: executedTestcases,
                htmlDatas: htmlReports);

            FinalReportGenerator.GenerateReport(baseSpec, reportType, carLine);
        }
    }
}
