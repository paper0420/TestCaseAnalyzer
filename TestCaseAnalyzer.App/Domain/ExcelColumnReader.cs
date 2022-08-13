using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TestCaseAnalyzer.App
{
    public class ExcelColumnReader
    {
        public ExcelColumnReader(IExcelDataReader reader, string file)
        {
            if (file == FileNames.KlhFile)
            {
                GetKLHFileIndex(reader);
                return;
            }

            if (file == FileNames.TestSpecFile)
            {
                GetTestspecFileIndex(reader);
                return;

            }

            throw new($"Unknown file name '{file}'.");
        }

        private void GetTestspecFileIndex(IExcelDataReader reader)
        {
            this.TestcaseSpecIDIndex = GetIndexNumber(reader, "Test Case ID");
            this.TestcaseSpecObjectiveIndex = GetIndexNumber(reader, "Test objective");
            this.TestcaseSpecRequirementIndex = GetIndexNumber(reader, "Current KLH");
            this.TestcaseSpecTypeIndex = GetIndexNumber(reader, "Type");
            this.TestcaseSpecG60Index = GetIndexNumber(reader, "G60");
            this.TestcaseSpecG70Index = GetIndexNumber(reader, "G70");
            this.TestcaseSpecClass1 = GetIndexNumber(reader, "Class1");
            this.TestcaseSpecClass2 = GetIndexNumber(reader, "Class2");
            this.TestcaseSpecClass3 = GetIndexNumber(reader, "Class3");
        }

        private void GetKLHFileIndex(IExcelDataReader reader)
        {
            this.KLHIDIndex = GetIndexNumber(reader, "Object ID from Original");
            this.KLHObjectiveIndex = GetIndexNumber(reader, "Englisch");
            this.ChangeStatusIndex = GetIndexNumber(reader, "A_Change Status");
            this.PanaStatusIndex = GetIndexNumber(reader, "A_Pana Status");

        }

        private static int GetIndexNumber(IExcelDataReader reader, string nameCheck)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var name = reader.GetValue(i)?.ToString();

                if (name == nameCheck)
                {
                    return i;
                }
            }

            throw new Exception($"Cannot find column '{nameCheck}'.");
        }

        public int TestcaseSpecIDIndex { get; set; }
        public int TestcaseSpecObjectiveIndex { get; set; }
        public int TestcaseSpecRequirementIndex { get; set; }
        public int TestcaseSpecTypeIndex { get; set; }
        public int TestcaseSpecG60Index { get; set; }
        public int TestcaseSpecG70Index { get; set; }
        public int TestcaseSpecClass1 { get; set; }
        public int TestcaseSpecClass2 { get; set; }
        public int TestcaseSpecClass3 { get; set; }
        public int KLHIDIndex { get; set; }
        public int KLHObjectiveIndex { get; set; }
        public int ChangeStatusIndex { get; set; }
        public int PanaStatusIndex { get; set; }
        public string VerificationMeasure { get; set; }
        public string Type { get; set; }



    }
}