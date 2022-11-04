using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.FileReader
{
    public static class ENG9HtmlReader
    {
        public static IEnumerable<HtmlData> ReadHtmlFullReport(string reportType,Dictionary<string,ENG9Testcase> testCases)
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

                if (!testCases.ContainsKey(expectedTestCaseId))
                {
                    continue;
                }

                var systemUnderTestNode = doc.DocumentNode.SelectSingleNode("//body/div")
                   .InnerText;

                if(systemUnderTestNode.Contains("System Under Test"))
                {
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
                        .Replace("_", ".")
                        .Remove(8, 1)
                        .Insert(8, "-");

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

                    var dsp = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                       .SelectSingleNode("table")
                       .SelectNodes("tr")[7]
                       .SelectNodes("td")[1]
                       .InnerText;

                    var pic = doc.DocumentNode.SelectNodes("//body/div/div")[1]
                      .SelectSingleNode("table")
                      .SelectNodes("tr")[0]
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

                    var testCaseID = $"{testCaseIDNode}";

                    if (processed.Contains(expectedTestCaseId))
                    {
                        Console.WriteLine($"Duplicate file detected {expectedTestCaseId}");
                        continue;
                    }

                    processed.Add(expectedTestCaseId);

                    var totalTestResultNode = doc.DocumentNode.SelectSingleNode("//body/center/table/tr/td")
                        .InnerText;

                    var totalTestResult = "";

                    if (totalTestResultNode.Contains("passed"))
                    {
                        totalTestResult = "OK";

                    }
                    else
                    {
                        totalTestResult = "NG";

                    }
                    var executedTestCaseNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                        .SelectNodes("div")[1]
                        .SelectNodes("table/tr")[0]
                        .SelectNodes("td")[1]
                        .InnerText
                        .ToInt();
                    var testCasePassedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                        .SelectNodes("div")[1]
                        .SelectNodes("table/tr")[1]
                        .SelectNodes("td")[1]
                        .InnerText
                        .ToInt();

                    var notExecutedTestCaseNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                        .SelectNodes("div")[1]
                        .SelectNodes("table/tr")[2]
                        .SelectNodes("td")[1]
                        .InnerText
                        .ToInt();

                    var testCaseFailedNumber = doc.DocumentNode.SelectNodes("//body/div")[1]
                      .SelectNodes("div")[1]
                      .SelectNodes("table/tr")[3]
                      .SelectNodes("td")[1]
                      .InnerText
                      .ToInt();

                    HtmlData data = new HtmlData();
                    data.ID = expectedTestCaseId;
                    data.TotalTestResult = totalTestResult;
                    data.NumberOfexecuted = executedTestCaseNumber;
                    data.NumberOfNotExecuted = notExecutedTestCaseNumber;
                    data.NumberOfPassed = testCasePassedNumber;
                    data.NumberOfFailed = testCaseFailedNumber;
                    data.HW = hw;
                    data.BTLD = btld;
                    data.SWFL = swfl;
                    data.SWFK = swfk;
                    data.DSP = dsp;
                    data.PIC = pic;
                    data.Location = location;

                    Console.WriteLine(expectedTestCaseId);
                    yield return data;


                }
                else
                {
                    Console.WriteLine($"No System under test section in this HTML. Tool will skip this TC {expectedTestCaseId}");
                    continue;

                }



                //if (expectedTestCaseId != testCaseID)
                //{
                //    Console.WriteLine($"File name {expectedTestCaseId} and HTML header {testCaseID} are mismatch");
                //    Console.WriteLine($"Window Login Name: {location}");
                //    testCaseID = expectedTestCaseId;

                //}

               
            }


        }

        private static string GetFileLabel(string file)
        {
            var file1 = Path.GetFileNameWithoutExtension(file)
            .SubstringFrom("[", inclusive: false)
            .SubstringUpTo("]")
            .Replace("-", "_");


            var fileName = $"{file1}";
            return fileName;
        }

        private static string[] GetPathName(string reportType)
        {

            string[] files = null;
            
            switch (reportType)
            {
                case "FTT":
                    files = Directory.GetFiles(FileNames.HtmlReportL1Folder, "*.html");

                    break;

                case "Fusa":
                    string[] htmlFiles1 = Directory.GetFiles(FileNames.HtmlReportL1Folder, "*.html");
                    string[] htmlFiles2 = Directory.GetFiles(FileNames.HtmlReportL2Folder, "*.html");
                    files = htmlFiles1.Union(htmlFiles2).ToArray();

                    break;

                case "Full":
                    files = Directory.GetFiles(FileNames.ENG9InputHTMLFolder, "*", SearchOption.AllDirectories);
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
