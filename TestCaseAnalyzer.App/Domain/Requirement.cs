using ExcelDataReader;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public class Requirement
    {
        public Requirement(IExcelDataReader reader, Header header)
        {
            this.changeStatus = reader.GetString(header.GetColumnIndex("A_Change Status"));
            this.panaStatus = reader.GetString(header.GetColumnIndex("A_Pana Status"));
            this.VerificationSpecStatus = reader.GetValue(header.GetColumnIndex("A_Verification_Specification_Status"))?.ToString();
            this.FusaType = reader.GetValue(header.GetColumnIndex("EAS_ASIL"))?.ToString();

            var idColumn = header.GetColumnIndex("Object ID from Original");
            
            var id = reader.GetValue(idColumn);
            
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                this.ID = reader.GetValue(idColumn).ToString();
            }

            this.Objective = reader.GetString(header.GetColumnIndex("Englisch"));
            this.VerificationMeasure = reader.GetValue(header.GetColumnIndex("A_Verification_Measure"))?.ToString()?.Replace("\n", "");
            
            //this.EpicIDs = reader
            //    .GetString(21)?
            //    .Split("\n", System.StringSplitOptions.RemoveEmptyEntries)
            //    ?? new string[0];
        }

        public string ID { get; }
        public string Objective { get; }
        public string changeStatus { get; }
        public string panaStatus { get; }
        public string VerificationSpecStatus { get; }
        public string VerificationMeasure { get; }
        public string Type { get; }
        public string[] EpicIDs { get; }
        public string FusaType { get; }



    }
}