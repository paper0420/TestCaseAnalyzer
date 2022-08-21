using IronXL;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App
{
    public class SpecParameters
    {
        public SpecParameters(
            List<Requirement> currentRequirments = null,
            List<Requirement> newRequirements = null,
            List<TestCaseOnlyExecutedItem> testCases = null,
            List<TestCase> panaTestCases = null,
            WorkSheet xlsSheet = null,
            List<HtmlData> htmlDatas = null,
            List<SafetyGoalKLH> safetyGoalKLHs = null)
        {
           
            this.NewRequirements = newRequirements;

            this.TestCases = testCases;
            this.TestCasesByID = testCases.ToDictionary(testCase=>testCase.ID);

            this.CurrentRequirements = currentRequirments;
            this.CurrentRequirementsByID = currentRequirments
                .Where(curRequirement => curRequirement.ID != null)
                .ToDictionary(curRequirement => curRequirement.ID);

            this.HtmlDatas = htmlDatas;
            this.HtmlDatasByID = htmlDatas
                .Where(htmlData => htmlData.ID != null)
                .DistinctBy(htmlData => htmlData.ID)
                .ToDictionary(htmlData => htmlData.ID);

            this.panaTestCases = panaTestCases;
            this.XlsSheet = xlsSheet;

            this.SafetyGoalKLHs = safetyGoalKLHs;
           

        }

        public List<Requirement> CurrentRequirements { get; set; }
        public List<Requirement> NewRequirements { get; set; }
        public List<TestCaseOnlyExecutedItem> TestCases { get; set; }
        public Dictionary<string, TestCaseOnlyExecutedItem> TestCasesByID { get; set; }
        public Dictionary<string, Requirement> CurrentRequirementsByID { get; set; }
        public List<HtmlData> HtmlDatas { get; set; }
        public Dictionary<string, HtmlData> HtmlDatasByID { get; set; }
        public List<TestCase> panaTestCases { get; set; }

        public WorkSheet XlsSheet { get; set; }
        public List<SafetyGoalKLH> SafetyGoalKLHs { get; set; }
    }
}
