using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class SYRitem
    {
        public SYRitem(IExcelDataReader reader)
        {
            this.ID = reader.GetValue(0)?.ToString();
            this.Objective = reader.GetString(1);
            this.Allocation = reader.GetValue(9)?.ToString();
            //this.TestResultFunctional = reader.GetString(188)?.Replace(" ", string.Empty);
            var idsAsString = reader.GetValue(3)?.ToString()?.Split('\n') ?? new string[0];

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
            this.SYR_ID = reader.GetValue(11)?.ToString();

        }

        public string ID { get; }
        public string Objective { get; }
        public string Allocation { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }
        public string SYR_ID { get; }

    }


}
