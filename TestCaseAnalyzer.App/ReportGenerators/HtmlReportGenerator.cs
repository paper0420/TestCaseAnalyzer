using System;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class HtmlReportGenerator : IReportGenerator
    {
        public void GenerateReport(SpecParameters spec)
        {
            Directory.CreateDirectory("report");

            string testcaseMenu = null;

            foreach(var testCase in spec.testCases)
            {
                if (testCase.ID != null)
                {
                    testcaseMenu += $"<a href='{testCase.ID.Replace("#", "")}.html'>{testCase.ID}</a><br>\n";
                }

            }

            foreach (var testCase in spec.testCases)
            {
                if (testCase.ID != null)
                {
                    string testCaseDetail = TestCaseHtmlGenerator.CreateTestCaseHtml(
                        spec.currentRequirements,testCase);

                    File.WriteAllText(
                        $"report/{testCase.ID.Replace("#", "")}.html",
                        testCaseDetail);
                }
            }



            File.WriteAllText("report/index.html", testcaseMenu);
        }
    }
}
