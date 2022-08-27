using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class HtmlReportGenerator : IReportGenerator
    {
        public void GenerateReport(SpecParameters spec)
        {
            Directory.CreateDirectory("report");

            string testcaseMenu = null;

            foreach(var testCase in spec.TestCases)
            {
                if (testCase.ID != null)
                {
                    testcaseMenu += $"<a href='{testCase.ID.Replace("#", "")}.html'>{testCase.ID}</a><br>\n";
                }

            }

            foreach (var testCase in spec.TestCases)
            {
                if (testCase.ID != null)
                {
                    string testCaseDetail = TestCaseHtmlGenerator.CreateTestCaseHtml(
                        spec.CurrentRequirements,testCase);

                    File.WriteAllText(
                        $"report/{testCase.ID.Replace("#", "")}.html",
                        testCaseDetail);
                }
            }



            File.WriteAllText("report/index.html", testcaseMenu);
        }
    }
}
