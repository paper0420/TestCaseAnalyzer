using ExcelDataReader;
using System.Collections.Generic;
using System.Linq;

namespace TestCaseAnalyzer.App
{
    public class TestCase
    {
        public TestCase(IExcelDataReader reader)
        {
            this.ID = reader.GetString(0);
            this.Objective = reader.GetString(8);

            var idsAsString = reader.GetValue(11)?.ToString()?.Split('\n') ?? new string[0];

            var requirementIds = new List<int>();
            var epicIds = new List<string>();

            foreach (var idAsString in idsAsString)
            {
                var isInt = int.TryParse(idAsString, out int idAsInt);
                
                if (isInt)
                {
                    requirementIds.Add(idAsInt);
                }
                else
                {
                    epicIds.Add(idAsString);
                }
            }

            this.RequirementIDs = requirementIds.ToArray();
            this.EpicIDs = epicIds.ToArray();
        }

        public string ID { get; }
        public string Objective { get; }
        public int[] RequirementIDs { get; }
        public string[] EpicIDs { get; }
    }
}
