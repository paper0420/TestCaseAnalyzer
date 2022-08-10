using IronXL;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class SpecParameters
    {
        public SpecParameters(
            List<Requirement> currentRequirments,
            List<Requirement> newRequirements,
            List<TestCaseOnlyExecutedItem> testCases,
            List<DeletedReq> delReqs,
            List<RejectedReq> rejReqs,
            List<SYRitem> syrItems,
            List<TestCase> panaTestCases,
            List<ENG9_Func_TestCase> eng9FuncTestCases,
            List<Dummy> syrLists,
            WorkSheet xlsSheet,
            List<LTestCases>[] allCarlineTestcaseDetails,
            List<string> allTestCaseIDs)
        {
            this.currentRequirements = currentRequirments;
            this.newRequirements = newRequirements;
            this.testCases = testCases;
            this.delReqs = delReqs;
            this.rejReqs = rejReqs;
            this.syrItems = syrItems;
            this.panaTestCases = panaTestCases;
            this.eng9FuncTestCases = eng9FuncTestCases;
            this.syrLists = syrLists;
            this.xlsSheet = xlsSheet;
            this.allCarlineTestcaseDetails = allCarlineTestcaseDetails;
            this.allTestCaseIDs = allTestCaseIDs;

        }

        public List<Requirement> currentRequirements { get; set; }
        public List<Requirement> newRequirements { get; set; }
        public List<TestCaseOnlyExecutedItem> testCases { get; set; }
        public List<DeletedReq> delReqs { get; set; }
        public List<RejectedReq> rejReqs { get; set; }
        public List<SYRitem> syrItems { get; set; }
        public List<TestCase> panaTestCases { get; set; }
        public List<ENG9_Func_TestCase> eng9FuncTestCases { get; set; }
        public List<Dummy> syrLists { get; set; }

        public WorkSheet xlsSheet { get; set; }
        public List<LTestCases>[] allCarlineTestcaseDetails { get; set; }

        public List<string> allTestCaseIDs { get; set; }
    }
}
