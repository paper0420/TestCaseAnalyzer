using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCaseAnalyzer.App
{
    public class FinalReportGenerationUI
    {
        public FinalReportGenerationUI()
        {
            List<string> carLineNames = new List<string> { "G60", "G70", "I20", "G26", "G28", "G08LCI", "U11" };
            Console.WriteLine("Please enter car line name");
            Console.WriteLine("G60 , G70, I20 , G26 , G28 , G08LCI , U11");
            Console.Write("Enter Car Line: ");
            this.CarLine = Console.ReadLine();

            while (!carLineNames.Contains(this.CarLine))
            {
                Console.WriteLine($"This car line {this.CarLine} is not available.");
                Console.Write("Enter Car Line: ");
                this.CarLine = Console.ReadLine();

            }

            List<string> reportNames = new List<string> { "HV", "Fusa", "Full" };
            Console.WriteLine("Please enter report type");
            Console.WriteLine("HV , Fusa , Full");
            Console.Write("Enter report type: ");
            this.ReportType = Console.ReadLine();
            while (!reportNames.Contains(this.ReportType))
            {
                Console.WriteLine($"This report type {this.ReportType} is not available.");
                Console.Write("Enter report type: ");
                this.ReportType = Console.ReadLine();

            }

            Console.WriteLine("Please enter SW release");
            this.SWRelease = Console.ReadLine();
            Console.WriteLine($"*****Car Line: {this.CarLine} SW Release: {this.SWRelease} Report type: {this.ReportType} *****");
            DateTime now = DateTime.Now;
            Console.WriteLine(now.ToString("F"));
        }


        public string CarLine { get; private set; }
        public string SWRelease { get; private set; }
        public  string ReportType { get; private set; }
    }
}
