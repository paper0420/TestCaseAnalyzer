using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class CarlineTestcases
    {
        public CarlineTestcases(IExcelDataReader reader)
        {


            var id = reader.GetValue(2);
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                this.ID = reader.GetString(2).Trim();

            }

            //DateTime dateTime;

            //this.TestTime = reader.GetValue(10)?.ToString();






            this.Category = reader.GetString(3);

        }



        public string ID { get; }
        public string TestTime { get; }
        public string Category { get; }
    }
}
