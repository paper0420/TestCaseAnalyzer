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

               string expectedTestCaseId = GetFileLabel(file);


                var hw = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                    .SelectSingleNode("table")
                    .SelectNodes("tr")[1]
                    .SelectNodes("td")[1]
                    .InnerText;

                var btld = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                    .SelectSingleNode("table")
                    .SelectNodes("tr")[4]
                    .SelectNodes("td")[1]
                    .InnerText
                    .Replace("_",".")
                    .Remove(8,1)
                    .Insert(8,"-");

                var swfl = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                    .SelectSingleNode("table")
                    .SelectNodes("tr")[5]
                    .SelectNodes("td")[1]
                    .InnerText
                    .Replace("_", ".")
                    .Remove(8, 1)
                    .Insert(8, "-");

                var swfk = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                    .SelectSingleNode("table")
                    .SelectNodes("tr")[6]
                    .SelectNodes("td")[1]
                    .InnerText;
                var apl = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                   .SelectSingleNode("table")
                   .SelectNodes("tr")[10]
                   .SelectNodes("td")[1]
                   .InnerText; 
                var dsp = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                   .SelectSingleNode("table")
                   .SelectNodes("tr")[11]
                   .SelectNodes("td")[1]
                   .InnerText
                   .Insert(3,".")
                   .Insert(1,".")
                   .Insert(0,"0");

                var pic = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                  .SelectSingleNode("table")
                  .SelectNodes("tr")[12]
                  .SelectNodes("td")[1]
                  .InnerText
                  .Insert(1, ".")
                  .Insert(0, "0.");

                var location = doc.DocumentNode.SelectNodes("//body/div/div")[3]
                  .SelectSingleNode("table")
                  .SelectSingleNode("tr")
                  .SelectNodes("td")[1]
                  .InnerText;

                var testCaseIDNode = doc.DocumentNode.SelectSingleNode("//body/table/tr/td/big")
                   .InnerText
                   .Replace("Report: ", "")
                   .Replace("_", "-");

                var testCaseID = $"#{testCaseIDNode}#";

                if (expectedTestCaseId != testCaseID)
                {
                    Console.WriteLine($"File name {expectedTestCaseId} and HTML header {testCaseID} are mismatch");
                    Console.WriteLine($"Window Login Name: {location}");
                    testCaseID = expectedTestCaseId;
                }

                if (processed.Contains(testCaseID))
                {
                    Console.WriteLine($"Duplicate file detected {testCaseID}");
                    continue;
                }

                processed.Add(testCaseID);

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
                data.HW = hw;
                data.BTLD = btld;
                data.SWFL = swfl;
                data.SWFK = swfk;
                data.APL = apl;
                data.DSP = dsp;
                data.PIC = pic;
                data.Location = location;


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
