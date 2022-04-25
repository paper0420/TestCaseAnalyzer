using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class TestCaseOnlyExecutedItem
    {
        public TestCaseOnlyExecutedItem(IExcelDataReader reader)
        {
            this.ID = reader.GetValue(0)?.ToString();
            var idsAsString = reader.GetValue(6)?.ToString()?.Split('\n') ?? new string[0];

            //var requirementIds = new List<int>();
            var requirementIds = new List<string>();
            var epicIds = new List<string>();

            foreach (var idAsString in idsAsString)
            {
                //var isInt = int.TryParse(idAsString, out int idAsInt);

                //if (isInt)
                //{
                //    requirementIds.Add(idAsInt);
                //}
                //else
                //{
                //    epicIds.Add(idAsString);
                //}

                requirementIds.Add(idAsString);


            }

            this.RequirementIDs = requirementIds.ToArray();
            this.EpicIDs = epicIds.ToArray();

        }

        public string ID { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }

    }
}
