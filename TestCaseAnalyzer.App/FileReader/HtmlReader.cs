using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.FileReader
{
    public static class HtmlReader
    {
        public static IEnumerable<HtmlData> ReadHtmlFullReport()
        {
            //string[][] all = new[]
            //{
            //    Directory.GetFiles(@"HtmlReportL3", "*.html"),
            //    Directory.GetFiles(@"HtmlReportL21", "*.html"),
            //    Directory.GetFiles(@"HtmlReportL21", "*.html")
            //};

            //string[] files = all.SelectMany(array => array).ToArray();

            string[] htmlFiles3 = Directory.GetFiles(@"HtmlReportL3", "*.html");
            string[] htmlFiles2 = Directory.GetFiles(@"HtmlReportL2", "*.html");
            string[] htmlFiles1 = Directory.GetFiles(@"HtmlReportL1", "*.html");

            var files = htmlFiles3.Union(htmlFiles2).Union(htmlFiles1).ToList();

            foreach (var file in files)
            {

                var path = file;

                var doc = new HtmlDocument();
                doc.Load(path);

                var testCaseIDNode = doc.DocumentNode.SelectSingleNode("//body/table/tr/td/big")
                    .InnerText
                    .Replace("Report: ", "")
                    .Replace("_", "-");
                var testCaseID = $"#{testCaseIDNode}#";

                var totalTestResultNode = doc.DocumentNode.SelectSingleNode("//body/center/table/tr/td")
                    .InnerText;

                var totalTestResult = "";

                if (totalTestResultNode.Contains("passed"))
                {
                    totalTestResult = "PASSED";
                }
                else
                {
                    totalTestResult = "FAILED";
                }

                var notExecutedTestCaseNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                    .SelectNodes("div")[1]
                    .SelectNodes("table/tr")[2]
                    .SelectNodes("td")[1]
                    .InnerText
                    .ToInt();

                var testCasePassedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                   .SelectNodes("div")[1]
                   .SelectNodes("table/tr")[3]
                   .SelectNodes("td")[1]
                   .InnerText
                   .ToInt();


                var testCaseFailedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                  .SelectNodes("div")[1]
                  .SelectNodes("table/tr")[4]
                  .SelectNodes("td")[1]
                  .InnerText
                  .ToInt();

                HtmlData data = new HtmlData();
                data.ID = testCaseID;
                data.TotalTestResult = totalTestResult;
                data.NumberOfNotExecuted = notExecutedTestCaseNumber;
                data.NumberOfPassed = testCasePassedNumber;
                data.NumberOfFailed = testCaseFailedNumber;


                yield return data;
            }
        }
    }
}
