using System;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    class Program
    {
        static void Main(string[] args)
        {
            App.RunMyApp(
                FileNames.KlhFile,
                "KLH10_14.xlsx");
        }
    
    }
}
