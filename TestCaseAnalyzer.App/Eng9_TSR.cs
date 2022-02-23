using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class Eng9_TSR
    {
        public Eng9_TSR(IExcelDataReader reader)
        {
            this.TSRID = reader.GetString(1);
            this.ASIL = reader.GetString(5);
            //Console.WriteLine(TSRID);
            //Console.WriteLine(ASIL);
         

        }

        public string TSRID { get; }
        public string ASIL { get; }



    }
}
