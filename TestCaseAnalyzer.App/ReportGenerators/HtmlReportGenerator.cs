using System;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class HtmlReportGenerator : IReportGenerator
    {
        public void GenerateReport(List<TestCase> testCases, List<Requirement> requirements)
        {
            Directory.CreateDirectory("report");
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
        }
    }
}
