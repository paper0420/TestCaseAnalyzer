using ExcelDataReader;
using System;

namespace TestCaseAnalyzer.App
{
    public class Eng9_TestCases
    {
  
        public Eng9_TestCases(IExcelDataReader reader)
        {
            
            var tsrID = reader.GetValue(2);

            if (!string.IsNullOrWhiteSpace(tsrID?.ToString()))
            {
                this.TSRID = reader.GetString(2).Trim();

            }

            var tcID = reader.GetValue(4);

            if (!string.IsNullOrWhiteSpace(tcID?.ToString()))
            {
                this.TestCaseID = reader.GetString(4).Trim();

            }
                
           //Console.WriteLine(TSRID);
            //Console.WriteLine(TestCaseID);

        }

        public string TSRID { get; }
        public string TestCaseID { get; }




    }
}