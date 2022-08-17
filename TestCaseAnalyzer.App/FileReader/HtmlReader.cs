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
        public static IEnumerable<HtmlData> ReadHtmlFullReport(string reportType)
        {
            var files = GetPathName(reportType);

            Console.WriteLine("***Collecting data from HTML reports**");

            var processed = new HashSet<string>();

            foreach (var file in files)
            {

                var path = file;

                var doc = new HtmlDocument();
                doc.Load(path);

                // H

                string expectedTestCaseId = GetFileLabel(file);

                var testCaseIDNode = doc.DocumentNode.SelectSingleNode("//body/table/tr/td/big")
                    .InnerText
                    .Replace("Report: ", "")
                    .Replace("_", "-");

                var testCaseID = $"#{testCaseIDNode}#";

                if (processed.Contains(testCaseID))
                {
                    Console.WriteLine($"Duplicate file detected {testCaseID}");
                    continue;
                }

                processed.Add(testCaseID);

                if (expectedTestCaseId != testCaseID)
                {
                    Console.WriteLine($"File {expectedTestCaseId} and test case {testCaseID} mismatch");
                }

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

        private static string GetFileLabel(string file)
        {
            var fileName1 = file
                .Replace("HtmlReportL2\\", "")
                .Replace(".html", "")
                .Replace("_report", "")
                .Replace("_", "-");
            
            var fileName2 = fileName1.Replace("HtmlReportL3\\", "");
            var fileName3 = fileName2.Replace("HtmlReportL1\\", "");
            var fileName = $"#{fileName3}#";

            return fileName;
        }

        private static string[] GetPathName(string reportType)
        {
            string[] files = null;

            switch (reportType)
            {
                case "HV":
                    files = Directory.GetFiles(@"HtmlReportL1", "*.html");

                    break;

                case "Fusa":
                    string[] htmlFiles1 = Directory.GetFiles(@"HtmlReportL1", "*.html");
                    string[] htmlFiles2 = Directory.GetFiles(@"HtmlReportL2", "*.html");
                    files = htmlFiles1.Union(htmlFiles2).ToArray();

                    break;

                case "Full":
                    string[] htmlFilesFull1 = Directory.GetFiles(@"HtmlReportL3", "*.html");
                    string[] htmlFilesFull2 = Directory.GetFiles(@"HtmlReportL2", "*.html");
                    string[] htmlFilesFull3 = Directory.GetFiles(@"HtmlReportL1", "*.html");

                    files = htmlFilesFull1.Union(htmlFilesFull2).Union(htmlFilesFull3).ToArray();
                    break;

                default:
                   
                    break;
            }
            //string[][] all = new[]
            //{
            //    Directory.GetFiles(@"HtmlReportL3", "*.html"),
            //    Directory.GetFiles(@"HtmlReportL21", "*.html"),
            //    Directory.GetFiles(@"HtmlReportL21", "*.html")
            //};

            //string[] files = all.SelectMany(array => array).ToArray();

            //if (reportType == "Full")
            //{
            //    string[] htmlFiles1 = Directory.GetFiles(@"HtmlReportL3", "*.html");
            //    string[] htmlFiles2 = Directory.GetFiles(@"HtmlReportL2", "*.html");
            //    string[] htmlFiles3 = Directory.GetFiles(@"HtmlReportL1", "*.html");
            //    files = htmlFiles1.Union(htmlFiles2).Union(htmlFiles3).ToArray();

            //}

            return files;
        }
    }
}
