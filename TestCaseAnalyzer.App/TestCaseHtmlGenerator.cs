using System.Linq;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class TestCaseHtmlGenerator
    {
        public static string CreateTestCaseHtml(
            List<Requirement> requirements, 
            TestCase testCase)
        {
            string requirementsHtml = GenerateRequirementsHtml(
                requirements,
                testCase);

            var epicsHtml = GenerateEpicsHtml(testCase);

            string testCaseDetail =
                $"<a href='./index.html'>Home</a>" +
                $"<h1>{testCase.ID}</h1>\n" +
                $"[{testCase.ItemClass1}],[{testCase.ItemClass2}],[{testCase.ItemClass3}]"+
                $"</br>"+
                $"Test Objective: {testCase.Objective}<br>\n" +
                $"<h2>Requirements</h2>" +
                $"{requirementsHtml}" +
                $"<h2>Epics</h2>" +
                $"{epicsHtml}";

            return testCaseDetail;
        }

        private static string GenerateEpicsHtml(TestCase testCase)
        {
            string epicHtml = "<ul>";

            foreach (var epicId in testCase.EpicIDs)
            {
                epicHtml += $"<li>{epicId}</li>";
            }

            epicHtml += "</ul>";

            return epicHtml;
        }

        private static string GenerateRequirementsHtml(List<Requirement> requirements, TestCase testCase)
        {
            var testCaseRequirements = requirements
                .Where(t => testCase.RequirementIDs.Any(c => c == t.ID))
                .ToList();

            string requirementsHtml = "<ul>";

            foreach (var requirement in testCaseRequirements)
            {
                var epicIds = requirement.EpicIDs
                    .Select(t => t.Replace("Epic(", "").TrimEnd(')'))
                    .Select(t => $"{t}")
                    .ToArray();

                var epicIdsAsString = epicIds.Any() 
                    ? $"<br>Epics: {string.Join(", ", epicIds)}"
                    : "";

                requirementsHtml += 
                    $"<li>" +
                    $"<strong>{requirement.ID}</strong> - [{requirement.changeStatus}] [{requirement.panaStatus}]"+
                    $"<br>{requirement.Objective}" +
                    $"{epicIdsAsString}" +
                    $"</li>\n";
            }

            requirementsHtml += "</ul>";
            
            return requirementsHtml;
        }
    }
}
