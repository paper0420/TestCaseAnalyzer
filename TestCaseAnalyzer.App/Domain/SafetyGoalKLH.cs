using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;

namespace TestCaseAnalyzer.App
{
    public class SafetyGoalKLH
    {
        public SafetyGoalKLH(IExcelDataReader reader, ExcelColumnReader index)
        {
            this.SG = reader.GetValue(index.SgIndex)?.ToString();
            this.ID = reader.GetValue(index.SGKLHID)?.ToString();

        }

        public string SG { get; }
        public string ID { get; }
        public string Attribute { get; }
        public string Asil { get; }
       



    }
}
