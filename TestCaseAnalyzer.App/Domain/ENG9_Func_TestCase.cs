using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class ENG9_Func_TestCase
    {
        public ENG9_Func_TestCase(IExcelDataReader reader)
        {
            this.ID = reader.GetValue(2)?.ToString();

            var syr_ID = reader.GetValue(0);
            if (!string.IsNullOrWhiteSpace(syr_ID?.ToString()))
            {
                this.SYR_ID = reader.GetValue(0) ?.ToString().Replace("#","");

            }

        }

        public string ID { get; }
        public string SYR_ID { get; }
    }
}
