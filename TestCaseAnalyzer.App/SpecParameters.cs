using IronXL;
using System.Collections.Generic;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App
{
    public class SpecParameters
    {
        public SpecParameters(
            List<Requirement> currentRequirments = null,
            List<Requirement> newRequirements = null,
            List<TestCaseOnlyExecutedItem> testCases = null,
            List<DeletedReq> delReqs = null ,
            List<RejectedReq> rejReqs = null,
            List<SYRitem> syrItems = null,
            List<TestCase> panaTestCases = null,
            List<ENG9_Func_TestCase> eng9FuncTestCases = null,
            List<Dummy> syrLists = null,
            WorkSheet xlsSheet = null,
            List<LTestCases>[] allCarlineTestcaseDetails = null,
            List<string> allTestCaseIDs = null,
            List<HtmlData> htmlDatas = null)
        {
            this.CurrentRequirements = currentRequirments;
            this.NewRequirements = newRequirements;
            this.TestCases = testCases;
            this.delReqs = delReqs;
            this.rejReqs = rejReqs;
            this.syrItems = syrItems;
            this.panaTestCases = panaTestCases;
            this.eng9FuncTestCases = eng9FuncTestCases;
            this.syrLists = syrLists;
            this.XlsSheet = xlsSheet;
            this.allCarlineTestcaseDetails = allCarlineTestcaseDetails;
            this.allTestCaseIDs = allTestCaseIDs;
            this.HtmlDatas = htmlDatas;

        }

        public List<Requirement> CurrentRequirements { get; set; }
        public List<Requirement> NewRequirements { get; set; }
        public List<TestCaseOnlyExecutedItem> TestCases { get; set; }
        public List<DeletedReq> delReqs { get; set; }
        public List<RejectedReq> rejReqs { get; set; }
        public List<SYRitem> syrItems { get; set; }
        public List<TestCase> panaTestCases { get; set; }
        public List<ENG9_Func_TestCase> eng9FuncTestCases { get; set; }
        public List<Dummy> syrLists { get; set; }

        public WorkSheet XlsSheet { get; set; }
        public List<LTestCases>[] allCarlineTestcaseDetails { get; set; }

        public List<string> allTestCaseIDs { get; set; }
        public List<HtmlData> HtmlDatas { get; set; }
    }
}
