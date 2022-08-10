using ExcelDataReader;

namespace TestCaseAnalyzer.App
{
    public class Requirement
    {
        public Requirement(IExcelDataReader reader)
        {
            int changeStatusIndex = ExcelColumnReader.GetColumnIndex("A_Change Status");

            var id = reader.GetValue(1);
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                this.ID = reader.GetValue(1)?.ToString();
            }


            this.Objective = reader.GetString(5);

            this.changeStatus = reader.GetString(2);
            this.panaStatus = reader.GetString(5);
            this.VerificationMeasure = reader.GetValue(13)?.ToString().Replace("\n","");
            this.Type = reader.GetValue(4)?.ToString();
            //this.EpicIDs = reader
            //    .GetString(21)?
            //    .Split("\n", System.StringSplitOptions.RemoveEmptyEntries)
            //    ?? new string[0];

        }

        public string ID { get; }
        public string Objective { get; }
        public string changeStatus { get; }
        public string panaStatus { get; }
        public string VerificationMeasure { get; }
        public string Type { get; }
        public string[] EpicIDs { get; }


    }
}