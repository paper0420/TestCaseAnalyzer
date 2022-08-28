using System;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class FinalReportGenerationUI
    {
        public static FinalReportGenerationUI ReadDataFromConsole(List<string> carLineNames)
        {
            Console.WriteLine("Please enter car line name");
            foreach(var carLine in carLineNames)
            {
                var car = carLine.Replace("#", "").Replace("#", "");
                Console.Write($"{car} ");
            }
            Console.WriteLine();
            Console.Write("Enter Car Line: ");
            
            var report = new FinalReportGenerationUI();
            report.CarLine = Console.ReadLine();
            report.CarLine = $"#{report.CarLine}#";

            while (!carLineNames.Contains(report.CarLine))
            {
                Console.WriteLine($"This car line {report.CarLine} is not available.");
                Console.Write("Enter Car Line: ");
                report.CarLine = Console.ReadLine();
                report.CarLine = $"#{report.CarLine}#";
            }

            List<string> reportNames = new List<string> { "HV", "Fusa", "Full" };
            Console.WriteLine("Please enter report type");
            Console.WriteLine("HV , Fusa , Full");
            Console.Write("Enter report type: ");
            report.ReportType = Console.ReadLine();
            while (!reportNames.Contains(report.ReportType))
            {
                Console.WriteLine($"This report type {report.ReportType} is not available.");
                Console.Write("Enter report type: ");
                report.ReportType = Console.ReadLine();
            }

            Console.WriteLine("Please enter SW release");
            report.SWRelease = Console.ReadLine();
            Console.WriteLine(
                $"*****Car Line: {report.CarLine} SW Release: {report.SWRelease} Report type: {report.ReportType} *****");
            DateTime now = DateTime.Now;
            Console.WriteLine(now.ToString("F"));

            return report;
        }

        public string CarLine { get; private set; }
        public string SWRelease { get; private set; }
        public  string ReportType { get; private set; }
    }
}
