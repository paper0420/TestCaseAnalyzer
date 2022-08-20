using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCaseAnalyzer.App.Domain
{
    public class HtmlData
    {
        public string ID { get; set; }
        public string TotalTestResult { get; set; }
        public int NumberOfPassed { get; set; }
        public int NumberOfFailed { get; set; }
        public int NumberOfNotExecuted { get; set; }



    }


}
