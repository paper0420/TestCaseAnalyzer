using ExcelDataReader;
using System;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class TestCase
    {
        public TestCase(IExcelDataReader reader)
        {
            this.ID = reader.GetString(0);
            this.Objective = reader.GetString(8);
            this.ItemClass1 = reader.GetString(4);
            this.ItemClass2 = reader.GetString(5);
            this.ItemClass3 = reader.GetString(6);
            this.TotalTestResults = reader.GetString(76)?.Replace(" ",string.Empty);
            this.TestResultFuSi = reader.GetString(135)?.Replace(" ",string.Empty);
            this.TestResultFunctional = reader.GetString(188)?.Replace(" ", string.Empty);
            this.VerificationMethod = reader.GetValue(25)?.ToString();
            this.TestCatHV = reader.GetValue(14)?.ToString().Replace("-", "");
            this.TestCatBasic = reader.GetValue(15)?.ToString().Replace("-", "");
            this.TestCatFusa = reader.GetValue(16)?.ToString().Replace("-", "");
            this.TestCatFunc = reader.GetValue(18)?.ToString().Replace("-", "");
            this.TestCatFull = reader.GetValue(19)?.ToString().Replace("-", "");

            Console.WriteLine($"{this.ID} {this.VerificationMethod} {this.TestCatHV}");

            var idsAsString = reader.GetValue(11)?.ToString()?.Split('\n') ?? new string[0];

            var requirementIds = new List<string>();
            var epicIds = new List<string>();

            foreach (var idAsString in idsAsString)
            {
                var isInt = int.TryParse(idAsString, out int idAsInt);
                
                if (isInt)
                {
                    requirementIds.Add(idAsString);
                }
                else
                {
                    epicIds.Add(idAsString);
                }
            }

            this.RequirementIDs = requirementIds.ToArray();
            this.EpicIDs = epicIds.ToArray();


            var varientG08 = reader.GetValue(29);
            if (!string.IsNullOrWhiteSpace(varientG08?.ToString()))
            {
                this.VariantG08 = reader.GetString(29).Trim();

            }

            var varientG26 = reader.GetValue(30);
            if (!string.IsNullOrWhiteSpace(varientG26?.ToString()))
            {
                this.VariantG26 = reader.GetString(30).Trim();

            }

            var varientG08LCI = reader.GetValue(31);
            if (!string.IsNullOrWhiteSpace(varientG08LCI?.ToString()))
            {
                this.VariantG08LCI = reader.GetString(31).Trim();

            }
            var varientGI20 = reader.GetValue(32);
            if (!string.IsNullOrWhiteSpace(varientGI20?.ToString()))
            {
                this.VariantI20 = reader.GetString(32).Trim();

            }

            HashSet<string> checkDupicate = new HashSet<string>();
            string comment = "";
            var comment1 = reader.GetValue(103)?.ToString();
            comment = $"{comment1}";
            checkDupicate.Add(comment1);

            var comment2 = reader.GetValue(119)?.ToString();
            if (!checkDupicate.Contains(comment2))
            {
                comment += $"\n{comment2}";
            }
            checkDupicate.Add(comment2);

            var comment3 = reader.GetValue(172)?.ToString();
            if (!checkDupicate.Contains(comment3))
            {
                comment += $"\n{comment3}";
            }
            checkDupicate.Add(comment3);


           
            this.Comment = comment;

        }

        public string ID { get; }
        public string Objective { get; }
        public string[] RequirementIDs { get; }
        public string[] EpicIDs { get; }

        public string VariantG08 { get; }
        public string VariantG26 { get; }
        public string VariantG08LCI { get; }

        public string VariantI20 { get; }

        public string ItemClass1 { get; }
        public string ItemClass2 { get; }
        public string ItemClass3 { get; }
        public string TotalTestResults { get; }
        public string TestResultFuSi { get; }
        public string TestResultFunctional { get; }
        public string VerificationMethod { get; }
        public string TestCatHV { get; }
        public string TestCatBasic { get; }
        public string TestCatFusa { get; }
        public string TestCatFunc { get; }
        public string TestCatFull { get; }
        public string Comment { get; }







    }
}
