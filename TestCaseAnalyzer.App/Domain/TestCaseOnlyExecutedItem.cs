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
            this.Objective = reader.GetString(4);
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
            this.G70 = reader.GetValue(13)?.ToString();
            this.Type = reader.GetValue(8)?.ToString();
            this.Result = reader.GetValue(17)?.ToString().Replace("\n"," ");

            this.ItemClass1 = reader.GetString(1);
            this.ItemClass2 = reader.GetString(2);
            this.ItemClass3 = reader.GetString(3);

        }

        public string ID { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }
        public string G70 { get; }
        public string Type { get; }
        public string Result { get; }
        public string Objective { get; }
        public string ItemClass1 { get; }
        public string ItemClass2 { get; }
        public string ItemClass3 { get; }



    }
}
