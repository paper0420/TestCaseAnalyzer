using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class TestCaseOnlyExecutedItem
    {
        public TestCaseOnlyExecutedItem(IExcelDataReader reader, ExcelColumnReader index)
        {
            this.ID = reader.GetValue(index.TestcaseSpecIDIndex)?.ToString();
            this.Objective = reader.GetString(index.TestcaseSpecObjectiveIndex);
            var idsAsString = reader.GetValue(index.TestcaseSpecRequirementIndex)?
                .ToString()?
                .Split('\n') ?? new string[0];

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
     


            this.G70 = GetCarline(reader, index.TestcaseSpecG70Index, "G70");
            this.G60 = GetCarline(reader, index.TestcaseSpecG60Index, "G60");
            this.G08LCI = GetCarline(reader, index.TestcaseSpecG08LCIIndex, "G08LCI");
            this.G26 = GetCarline(reader, index.TestcaseSpecG26Index, "G26");
            this.G28 = GetCarline(reader, index.TestcaseSpecG28Index, "G28");
            this.I20 = GetCarline(reader, index.TestcaseSpecI20Index, "I20");
            this.U11 = GetCarline(reader, index.TestcaseSpecU11Index, "U11");


            this.Carline = $"{this.G70},{this.G60}, {this.G08LCI},{this.G26}, {this.G28}, {this.I20}, {this.U11}";

            this.Type = reader.GetValue(index.TestcaseSpecTypeIndex)?.ToString();
            this.Result = reader.GetValue(index.TestcaseSpecResultIndex)?.ToString().Replace("\n", " ");

            this.ItemClass1 = reader.GetString(index.TestcaseSpecClass1Index);
            this.ItemClass2 = reader.GetString(index.TestcaseSpecClass2Index);
            this.ItemClass3 = reader.GetString(index.TestcaseSpecClass3Index);

            this.Comment = reader.GetValue(index.TestcaseSpecCommentIndex)?.ToString();




        }

        private static string GetCarline(IExcelDataReader reader, int index,string carlineName)
        {
            var checkCarline = reader.GetValue(index)?.ToString();
            if(checkCarline != null)
            {
                if (checkCarline.Contains("X"))
                {
                    checkCarline = carlineName;
                    return checkCarline;
                }

            }

            return "No";
        }

        public string ID { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }
        public string G70 { get; }
        public string G60 { get; }
        public string G08LCI { get; }
        public string G26 { get; }
        public string G28 { get; }
        public string I20 { get; }
        public string U11 { get; }
        public string Type { get; }
        public string Result { get; }
        public string Objective { get; }
        public string ItemClass1 { get; }
        public string ItemClass2 { get; }
        public string ItemClass3 { get; }
        public string Carline { get; }
        public string Comment { get; }




    }
}
