using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;

namespace TestCaseAnalyzer.App
{
    public static class HtmlReader
    {
        public static void ReadHtmlFullReport()
        {
            var path = @"HtmlReportL3\SYQT-04-0003-0028_report.html";

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
                totalTestResult = "passed";
            }
            else
            {
                totalTestResult = "failed";
            }

            var notExecutedTestCaseNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                .SelectNodes("div")[1]
                .SelectNodes("table/tr")[2]
                .SelectNodes("td")[1];

            var testCasePassedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
               .SelectNodes("div")[1]
               .SelectNodes("table/tr")[3]
               .SelectNodes("td")[1];

            var testCaseFailedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
              .SelectNodes("div")[1]
              .SelectNodes("table/tr")[3]
              .SelectNodes("td")[1];

            Console.WriteLine(testCaseID);
            Console.WriteLine(totalTestResult);
            Console.WriteLine(notExecutedTestCaseNumber.InnerText);
            Console.WriteLine(testCasePassedNumber.InnerText);
            Console.WriteLine(testCaseFailedNumber.InnerText);



            



        }
    }
}
