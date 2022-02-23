using IronXL;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class SpecParameters
    {
        public SpecParameters(
            List<Requirement> requirements,
            List<TestCase> testCases,
            List<DeletedReq> delReqs,
            List<RejectedReq> rejReqs,
            WorkSheet xlsSheet,
            List<LTestCases>[] allCarlineTestcaseDetails,
            List<string> allTestCaseIDs)
        {
            this.requirements = requirements;
            this.testCases = testCases;
            this.delReqs = delReqs;
            this.rejReqs = rejReqs;
            this.xlsSheet = xlsSheet;
            this.allCarlineTestcaseDetails = allCarlineTestcaseDetails;
            this.allTestCaseIDs = allTestCaseIDs;

        }

        public List<Requirement> requirements { get; set; }
        public List<TestCase> testCases { get; set; }
        public List<DeletedReq> delReqs { get; set; }
        public List<RejectedReq> rejReqs { get; set; }
        public WorkSheet xlsSheet { get; set; }
        public List<LTestCases>[] allCarlineTestcaseDetails { get; set; }

        public List<string> allTestCaseIDs { get; set; }
    }
}
