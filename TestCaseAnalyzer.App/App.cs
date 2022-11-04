using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestCaseAnalyzer.App.Database;
using TestCaseAnalyzer.App.Domain;
using TestCaseAnalyzer.App.FileReader;
using TestCaseAnalyzer.App.ReportGenerators;
using TestCaseAnalyzer.App.Spec;

namespace TestCaseAnalyzer.App
{
    public class App
    {
        public static void RunMyApp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            
            //var currentKLHs = ExcelTableReader.ReadFile(FileNames.TestSpecFile, "KLH", (t, y) => Requirement.CreateOrNull(t, y)).DataRows;
            //var safetyGoalKLHs = ExcelTableReader.ReadFile(FileNames.TestSpecFile, "SafetyGoal", (t, y) => SafetyGoalKLH.CreateOrNull(t, y)).DataRows;
            
            //var testItemWorksheet = ExcelTableReader.ReadFile(FileNames.TestSpecFile, "Test_Item", (t, y) => TestCaseOnlyExecutedItem.CreateOrNull(t, y));
            //var executedTestcases = testItemWorksheet.DataRows.ToList();

            var eng9TestCases = ExcelTableReader.ReadFile(FileNames.ENG9TestSpec, "Test_Item", (t, y) => ENG9Testcase.CreateOrNull(t, y)).DataRows;

            //var panaTestCases = reader.ReadPanaFile("data.xlsm", "5.テスト項目Test item", 18, t => new TestCase(t)).ToList();
            //var spec = new SpecParameters(panaTestCases: panaTestCases, testCases: executedTestcases);
            //TestSpecExcelGenerator.UpdateTestSpecExcel(spec);

            //var userInput = FinalReportGenerationUI.ReadDataFromConsole(CarLineNames.carLineNames);
            //var htmlReports = HtmlReader.ReadHtmlFullReport(userInput.ReportType).ToList();

            //var baseSpec = new SpecParameters(
            //    currentRequirments: currentKLHs,
            //    testCases: executedTestcases,
            //    htmlDatas: htmlReports,
            //    safetyGoalKLHs: safetyGoalKLHs);

            //FinalReportGenerator.GenerateReport(baseSpec, userInput.ReportType, userInput.CarLine, userInput.SWRelease);
            Dictionary<string,ENG9Testcase> testCasesByID = eng9TestCases.ToDictionary(testCase => testCase.ID);

            var eng9HtmlReports = ENG9HtmlReader.ReadHtmlFullReport("Full", testCasesByID).ToList();   
            
            var eng9Spec = new ENG9Spec(
            testCases: eng9TestCases,
            htmlDatas: eng9HtmlReports);
            ENG9FinalReportGenerator.ENG9GenerateReport(eng9Spec,"Full","G26","00");
        }




    }
}
