using ExcelDataReader;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public class TestCaseOnlyExecutedItem
    {
        public TestCaseOnlyExecutedItem(IExcelDataReader reader, Header header)
        {
            this.ID = reader.GetValue(header.GetColumnIndex("Test Case ID"))?.ToString();
            this.Objective = reader.GetString(header.GetColumnIndex("Test objective"));
            var idsAsString = reader.GetValue(header.GetColumnIndex("Current KLH"))?
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

            this.Carlines = GetAvailableCarlines(
                    reader,
                    header,
                    CarLineNames.carLineNames)
                .ToList();
            
            this.Type = reader.GetValue(header.GetColumnIndex("Type"))?.ToString();
            this.Result = reader.GetValue(header.GetColumnIndex("Result"))?.ToString()?.Replace("\n", " ");

            this.ItemClass1 = reader.GetString(header.GetColumnIndex("Class1"));
            this.ItemClass2 = reader.GetString(header.GetColumnIndex("Class2"));
            this.ItemClass3 = reader.GetString(header.GetColumnIndex("Class3"));

            this.Comment = reader.GetValue(header.GetColumnIndex("Comment"))?.ToString();

            this.VerificationMethod = reader.GetValue(header.GetColumnIndex("Verification Method"))?.ToString();
            this.TestCatHV = reader.GetValue(header.GetColumnIndex("TestCat-HV"))?.ToString();
            this.TestCatBasic = reader.GetValue(header.GetColumnIndex("TestCat-HV"))?.ToString();
            this.TestCatFusa = reader.GetValue(header.GetColumnIndex("TestCat-Fusa"))?.ToString();
            this.TestCatFunc = reader.GetValue(header.GetColumnIndex("TestCat-Func"))?.ToString();
            this.TestCatFull = reader.GetValue(header.GetColumnIndex("TestCat-Full"))?.ToString();
        }

        public IList<string> Carlines { get; }

        private static IEnumerable<string> GetAvailableCarlines(
            IExcelDataReader reader,
            Header header,
            List<string> carLineNames)
        {
            foreach (var carLineName in carLineNames)
            {
                var columnIndex = header.GetColumnIndex(carLineName);
            
                var contains = reader.GetValue(columnIndex)?.ToString()?.Contains("X") == true;

                if (contains)
                {
                    yield return carLineName;
                }
            }
        }

        public string ID { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }
        
        public string Type { get; }
        public string Result { get; }
        public string Objective { get; }
        public string ItemClass1 { get; }
        public string ItemClass2 { get; }
        public string ItemClass3 { get; }
        public string Comment { get; }
        public string VerificationMethod { get; }
        public string TestCatHV { get; }
        public string TestCatBasic { get; }
        public string TestCatFusa { get; }
        public string TestCatFunc { get; }
        public string TestCatFull { get; }
    }
}
