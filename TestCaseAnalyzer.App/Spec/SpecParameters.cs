using IronXL;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.Spec
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

            NewRequirements = newRequirements;

            TestCases = testCases;
            TestCasesByID = testCases.ToDictionary(testCase => testCase.ID);

            CurrentRequirements = currentRequirments;
            CurrentRequirementsByID = currentRequirments
                .Where(curRequirement => curRequirement.ID != null)
                .ToDictionary(curRequirement => curRequirement.ID);

            HtmlDatas = htmlDatas;

            if (htmlDatas != null)
            {
                HtmlDatasByID = htmlDatas
                .Where(htmlData => htmlData.ID != null)
                .DistinctBy(htmlData => htmlData.ID)
                .ToDictionary(htmlData => htmlData.ID);

            }


            this.panaTestCases = panaTestCases;
            XlsSheet = xlsSheet;

            SafetyGoalKLHs = safetyGoalKLHs;


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
