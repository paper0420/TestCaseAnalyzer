using ClosedXML.Excel;
using IronXL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestCaseAnalyzer.App.FileReader;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    public class App
    {
        public static void RunMyApp(string newFile, string currentFile)
        {
            var wb = new XLWorkbook("tp.xlsx");
            var ws = wb.Worksheet("FuSi");
            ws.Cell("A1").Value = "AAAA";
            wb.Save();
            return;



            List<string> carLineNames = new List<string>{"G60","G70","I20","G26","G28", "G08LCI", "U11" };
            Console.WriteLine("Please enter car line name");
            Console.WriteLine("G60 , G70, I20 , G26 , G28 , G08LCI , U11");
            Console.Write("Enter Car Line: ");
            string carLine = Console.ReadLine();

            while (!carLineNames.Contains(carLine))
            {
                Console.WriteLine($"This car line {carLine} is not available.");
                Console.Write("Enter Car Line: ");
                carLine = Console.ReadLine();

            }

            List<string> reportNames = new List<string> { "HV", "Fusa", "Full"};
            Console.WriteLine("Please enter report type");
            Console.WriteLine("HV , Fusa , Full");
            Console.Write("Enter report type: ");
            string reportType = Console.ReadLine();
            while (!reportNames.Contains(reportType))
            {
                Console.WriteLine($"This report type {reportType} is not available.");
                Console.Write("Enter report type: ");
                reportType = Console.ReadLine();

            }

            Console.WriteLine("Please enter SW release");
            string swRelease = Console.ReadLine();

            Console.WriteLine($"*****Car Line: {carLine} SW Release: {swRelease} Report type: {reportType} *****");


            DateTime now = DateTime.Now;
            Console.WriteLine(now.ToString("F"));
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new ExcelReader();
            var currentKLHs = reader.ReadFile(FileNames.TestSpecFile, "KLH_BL11.1", 1, (t,y) => new Requirement(t,y)).ToList();
            var safetyGoalKLHs = reader.ReadFile(FileNames.TestSpecFile, "SafetyGoal",1,(t,y) => new SafetyGoalKLH(t,y)).ToList();
            var executedTestcases = reader.ReadFile(FileNames.TestSpecFile, "Test_Item", 1, (t, y) => new TestCaseOnlyExecutedItem(t, y)).ToList();
            var htmlReports = HtmlReader.ReadHtmlFullReport(reportType).ToList();

            //var panaTestCases = reader.ReadPanaFile("data.xlsm", "5.テスト項目Test item", 18, t => new TestCase(t)).ToList();
            //var spec = new SpecParameters(panaTestCases: panaTestCases, testCases: executedTestcases);
            //TestSpecExcelGenerator.UpdateTestSpecExcel(spec);


            var baseSpec = new SpecParameters(
                currentRequirments: currentKLHs,
                testCases: executedTestcases,
                htmlDatas: htmlReports,
                safetyGoalKLHs: safetyGoalKLHs);

            FinalReportGenerator.GenerateReport(baseSpec, reportType, carLine,swRelease);
        }







    }
}
