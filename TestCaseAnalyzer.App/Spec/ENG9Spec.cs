using IronXL;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.Spec
{
    public class ENG9Spec
    {
        public ENG9Spec(
            List<Requirement> currentRequirments = null,
            List<ENG9Testcase> testCases = null,
            WorkSheet xlsSheet = null,
            List<HtmlData> htmlDatas = null)
        {


            TestCases = testCases;
            TestCasesByID = testCases.ToDictionary(testCase => testCase.ID);

            //CurrentRequirements = currentRequirments;
            //CurrentRequirementsByID = currentRequirments
            //    .Where(curRequirement => curRequirement.ID != null)
            //    .ToDictionary(curRequirement => curRequirement.ID);

            HtmlDatas = htmlDatas;

            if (htmlDatas != null)
            {
                HtmlDatasByID = htmlDatas
                .Where(htmlData => htmlData.ID != null)
                .DistinctBy(htmlData => htmlData.ID)
                .ToDictionary(htmlData => htmlData.ID);

            }

            //XlsSheet = xlsSheet;


        }

        public List<Requirement> CurrentRequirements { get; set; }
        public List<ENG9Testcase> TestCases { get; set; }
        public Dictionary<string, ENG9Testcase> TestCasesByID { get; set; }
        public Dictionary<string, Requirement> CurrentRequirementsByID { get; set; }
        public List<HtmlData> HtmlDatas { get; set; }
        public Dictionary<string, HtmlData> HtmlDatasByID { get; set; }

        public WorkSheet XlsSheet { get; set; }
    }
}
