using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TestCaseAnalyzer.App
{
    public class ExcelColumnReader
    {
        public ExcelColumnReader(IExcelDataReader reader, string file, string sheet)
        {
            if (file == FileNames.KlhFile)
            {
                GetKLHFileIndex(reader);
                return;
            }

            if (file == FileNames.TestSpecFile)
            {
                switch (sheet)
                {
                    case "Test_Item":
                        GetTestspecFileIndex(reader);
                        break;

                    case "SafetyGoal":
                        GetSafetyGoalSheetIndex(reader);
                        break;

                    case "KLH_BL11.1":
                        GetKLHFileIndex(reader);
                        break;
                        default:
                        throw new($"Unknown sheet name '{sheet}'.");

                }
                
                return;

            }

            throw new($"Unknown file name '{file}'.");
        }

        private void GetSafetyGoalSheetIndex(IExcelDataReader reader)
        {
            this.SGKLHID = GetIndexNumber(reader, "Object ID from Original");
            this.AttributeIndex = GetIndexNumber(reader, "attribute");
            this.EAS_ASILIndex = GetIndexNumber(reader, "EAS_ASIL");
            this.SgIndex = GetIndexNumber(reader, "SG");

        }

        private void GetTestspecFileIndex(IExcelDataReader reader)
        {
            this.TestcaseSpecIDIndex = GetIndexNumber(reader, "Test Case ID");
            this.TestcaseSpecObjectiveIndex = GetIndexNumber(reader, "Test objective");
            this.TestcaseSpecRequirementIndex = GetIndexNumber(reader, "Current KLH");
            this.TestcaseSpecTypeIndex = GetIndexNumber(reader, "Type");
            this.TestcaseSpecG60Index = GetIndexNumber(reader, "G60");
            this.TestcaseSpecG70Index = GetIndexNumber(reader, "G70");
            this.TestcaseSpecG08LCIIndex = GetIndexNumber(reader, "G08LCI");
            this.TestcaseSpecG26Index = GetIndexNumber(reader, "G26");
            this.TestcaseSpecG28Index = GetIndexNumber(reader, "G28");
            this.TestcaseSpecI20Index = GetIndexNumber(reader, "I20");
            this.TestcaseSpecU11Index = GetIndexNumber(reader, "U11");
            this.TestcaseSpecClass1Index = GetIndexNumber(reader, "Class1");
            this.TestcaseSpecClass2Index = GetIndexNumber(reader, "Class2");
            this.TestcaseSpecClass3Index = GetIndexNumber(reader, "Class3");
            this.TestcaseSpecResultIndex = GetIndexNumber(reader, "Result");
            this.TestcaseSpecCommentIndex = GetIndexNumber(reader, "Comment");
            this.VerificationMethodIndex = GetIndexNumber(reader, "Verification Method");
            this.TestCatHVIndex = GetIndexNumber(reader, "TestCat-HV");
            this.TestCatBasicIndex = GetIndexNumber(reader, "TestCat-Basic");
            this.TestCatFusaIndex = GetIndexNumber(reader, "TestCat-Fusa");
            this.TestCatFuncIndex = GetIndexNumber(reader, "TestCat-Func");
            this.TestCatFullIndex = GetIndexNumber(reader, "TestCat-Full");



        }

        private void GetKLHFileIndex(IExcelDataReader reader)
        {
            this.KLHIDIndex = GetIndexNumber(reader, "Object ID from Original");
            this.KLHObjectiveIndex = GetIndexNumber(reader, "Englisch");
            this.ChangeStatusIndex = GetIndexNumber(reader, "A_Change Status");
            this.PanaStatusIndex = GetIndexNumber(reader, "A_Pana Status");
            this.KLHEAS_ASILIndex = GetIndexNumber(reader, "EAS_ASIL");
            this.VerificationSpecStatusIndex = GetIndexNumber(reader, "A_Verification_Specification_Status");

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
        public int TestcaseSpecG08LCIIndex { get; set; }
        public int TestcaseSpecG26Index { get;set; }
        public int TestcaseSpecG28Index { get; set; }
        public int TestcaseSpecI20Index { get; set; }
        public int TestcaseSpecU11Index { get; set; }
        public int TestcaseSpecClass1Index { get; set; }
        public int TestcaseSpecClass2Index { get; set; }
        public int TestcaseSpecClass3Index { get; set; }
        public int TestcaseSpecResultIndex { get;set; }
        public int TestcaseSpecCommentIndex { get; set; }
        public int VerificationMethodIndex { get; private set; }
        public int TestCatHVIndex { get; private set; }
        public int TestCatBasicIndex { get; private set; }
        public int TestCatFusaIndex { get; private set; }
        public int TestCatFuncIndex { get; private set; }
        public int TestCatFullIndex { get; private set; }
        public int KLHIDIndex { get; set; }
        public int KLHObjectiveIndex { get; set; }
        public int ChangeStatusIndex { get; set; }
        public int PanaStatusIndex { get; set; }
        public int KLHEAS_ASILIndex { get; set; }
        public int VerificationSpecStatusIndex { get; private set; }
        public string VerificationMeasure { get; set; }
        public string Type { get; set; }
        public int SGKLHID { get; private set; }
        public int AttributeIndex { get; private set; }
        public int EAS_ASILIndex { get; private set; }
        public int SgIndex { get; private set; }
    }
}