using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.FileReader;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    public class App
    {
        public static void RunMyApp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new ExcelReader();

            var currentKLHs = reader
                .ReadFile(FileNames.TestSpecFile, "KLH", (t, y) => new Requirement(t, y))
                .DataRows;
            
            
            var safetyGoalKLHs = reader.ReadFile(FileNames.TestSpecFile, "SafetyGoal", (t, y) => new SafetyGoalKLH(t, y)).DataRows.ToList();
            var executedTestcases = reader.ReadFile(FileNames.TestSpecFile, "Test_Item", (t, y) => new TestCaseOnlyExecutedItem(t, y)).DataRows.ToList();

            //var panaTestCases = reader.ReadPanaFile("data.xlsm", "5.テスト項目Test item", 18, t => new TestCase(t)).ToList();
            //var spec = new SpecParameters(panaTestCases: panaTestCases, testCases: executedTestcases);
            //TestSpecExcelGenerator.UpdateTestSpecExcel(spec);
            
            var lineNames = new List<string> { "G60", "G70", "I20", "G26", "G28", "G08LCI", "U11" };

            var userInput = FinalReportGenerationUI.ReadDataFromConsole(lineNames);
            var htmlReports = HtmlReader.ReadHtmlFullReport(userInput.ReportType).ToList();

            var baseSpec = new SpecParameters(
                currentRequirments: currentKLHs,
                testCases: executedTestcases,
                htmlDatas: htmlReports,
                safetyGoalKLHs: safetyGoalKLHs);

            FinalReportGenerator.GenerateReport(baseSpec, userInput.ReportType, userInput.CarLine, userInput.SWRelease);
        }




    }
}
