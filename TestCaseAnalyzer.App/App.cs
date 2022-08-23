using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestCaseAnalyzer.App.FileReader;
using TestCaseAnalyzer.App.ReportGenerators;
using TestCaseAnalyzer.App.Analysis;

namespace TestCaseAnalyzer.App
{
    public class App
    {
        public static void RunMyApp(string newFile, string currentFile)
        {
            var userInput = new FinalReportGenerationUI();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new ExcelReader();
            var currentKLHs = reader.ReadFile(FileNames.TestSpecFile, "KLH_BL11.1", 1, (t,y) => new Requirement(t,y)).ToList();
            var safetyGoalKLHs = reader.ReadFile(FileNames.TestSpecFile, "SafetyGoal",1,(t,y) => new SafetyGoalKLH(t,y)).ToList();
            var executedTestcases = reader.ReadFile(FileNames.TestSpecFile, "Test_Item", 1, (t, y) => new TestCaseOnlyExecutedItem(t, y)).ToList();

            //var panaTestCases = reader.ReadPanaFile("data.xlsm", "5.テスト項目Test item", 18, t => new TestCase(t)).ToList();
            //var spec = new SpecParameters(panaTestCases: panaTestCases, testCases: executedTestcases);
            //TestSpecExcelGenerator.UpdateTestSpecExcel(spec);
           
            var htmlReports = HtmlReader.ReadHtmlFullReport(userInput.ReportType).ToList();

            var baseSpec = new SpecParameters(
                currentRequirments: currentKLHs,
                testCases: executedTestcases,
                htmlDatas: htmlReports,
                safetyGoalKLHs: safetyGoalKLHs);


            //FindTestCasewithWrongVerificationLinked.FindTCwWrongVeriMea(baseSpec);
            FinalReportGenerator.GenerateReport(baseSpec, userInput.ReportType, userInput.CarLine,userInput.SWRelease);
        }




    }
}
