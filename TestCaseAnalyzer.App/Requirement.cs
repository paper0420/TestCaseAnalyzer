using ExcelDataReader;

namespace TestCaseAnalyzer.App
{
    public class Requirement
    {
        public Requirement(IExcelDataReader reader)
        {

            var id = reader.GetValue(1);
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                this.ID = (int)reader.GetDouble(1);
            }


            this.Objective = reader.GetString(5);
            //this.EpicIDs = reader
            //    .GetString(21)?
            //    .Split("\n", System.StringSplitOptions.RemoveEmptyEntries)
            //    ?? new string[0];

            this.changeStatus = reader.GetString(3);
            this.panaStatus = reader.GetString(10);
            
        }

        public int ID { get; }
        public string Objective { get; }
        public string[] EpicIDs { get; }
        public string changeStatus { get; }
        public string panaStatus { get; }

    }
}