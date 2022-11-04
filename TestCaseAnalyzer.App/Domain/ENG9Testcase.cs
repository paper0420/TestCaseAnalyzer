using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App.Domain
{
    public class ENG9Testcase
    {
        internal static ENG9Testcase? CreateOrNull(IExcelDataReader reader, Header header)
        {
            var id = reader.GetStringOrNull(header.GetColumnIndex("Test Case ID"))?.ToString();


            if (id == null)
            {
                return null;
            }

            var result = new ENG9Testcase();

            result.ID = id;
            result.TSRID = reader.GetStringOrNull(header.GetColumnIndex("TSR ID"))?.ToString()?
                .Trim()
                .Split('\n')?
                .ToList();
            result.SYRID = reader.GetStringOrNull(header.GetColumnIndex("SYR ID"))?.ToString()?
                .Trim()
                .Split('\n')?
                .ToList();

            result.KLHID = reader.GetStringOrNull(header.GetColumnIndex("KLH ID"))?.ToString()?
                .Trim()
                .Split('\n')?
                .ToList();
            result.ICSID = reader.GetStringOrNull(header.GetColumnIndex("ICS ID"))?.ToString()?
                .Trim()
                .Split('\n')?
                .ToList();

            result.Objective = reader.GetStringOrNull(header.GetColumnIndex("Objective"))?.ToString();
            result.Catagory = reader.GetStringOrNull(header.GetColumnIndex("Catagory"))?.ToString();
            result.FusaType = reader.GetStringOrNull(header.GetColumnIndex("FusaType"))?.ToString();
            result.Carlines = reader.GetStringOrNull(header.GetColumnIndex("Available Car lines"))?.ToString()?
                .Split('\n').ToList();

            result.RelatedTSRID = reader.GetStringOrNull(header.GetColumnIndex("Related TSR"))?.ToString();
            result.SafetyGoal = reader.GetStringOrNull(header.GetColumnIndex("Safety Goal"))?.ToString();
            result.SubIDs = reader.GetStringOrNull(header.GetColumnIndex("Sub Test Case ID"))?.ToString()?
                   .Trim()
                   .Split('\n')?
                   .ToList();

            result.FunctionCatagory = reader.GetStringOrNull(header.GetColumnIndex("Function Catagory"))?.ToString();
            result.ErrorFactor = reader.GetStringOrNull(header.GetColumnIndex("Error Factor"))?.ToString();




            return result;
        }

        public List<string> REQID { get; set; } = new List<string> { };
        public List<string> KLHID { get; set; } = new List<string> { };
        public List<string> ICSID { get; set; } = new List<string> { };
        public List<string> TSRID { get; set; } = new List<string> { };
        public List<string> SYRID { get; set; } = new List<string> { };
        public string Objective { get; set; }
        public List<string> Carlines { get; set; } = new List<string> { };
        public string FusaType { get; set; }
        public string Catagory { get; set; }
        public string ID { get; set; }
        public string RelatedTSRID { get; set; }
        public string SafetyGoal { get; set; }
        public List<string> SubIDs { get; set; }
        public string FunctionCatagory { get; set; }
        public string ErrorFactor { get; set; }


    }
}
