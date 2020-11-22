using ExcelDataReader;

namespace TestCaseAnalyzer.App
{
    public class Requirement
    {
        public Requirement(IExcelDataReader reader)
        {
            this.ID = (int)reader.GetDouble(0);
            this.Objective = reader.GetString(4);
            this.EpicIDs = reader
                .GetString(21)?
                .Split("\n", System.StringSplitOptions.RemoveEmptyEntries) 
                ?? new string[0];
        }

        public int ID { get; }
        public string Objective { get;}
        public string[] EpicIDs { get;}
    }
}