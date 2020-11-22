using IronXL;
using System.Collections.Generic;
using System.Linq;

namespace TestCaseAnalyzer.App
{
    public class TestCaseExcelGenerator
    {
        public static void CreateTestCaseExcel(
            List<Requirement> requirements,
            TestCase testCase,
            WorkSheet xlsSheet)
        {
            GenerateRequirementsExcel(
                requirements,
                testCase,
                xlsSheet);

            xlsSheet["A1"].Value = "Test Case ID";
            xlsSheet["C1"].Value = $"{testCase.ID}";
            xlsSheet["A2"].Value = "Test Objective";
            xlsSheet["C2"].Value = $"{testCase.Objective}";
            xlsSheet["A3"].Value = "Requirements";
        }

        private static void GenerateRequirementsExcel(List<Requirement> requirements, TestCase testCase, WorkSheet xlsSheet)
        {
            var testCaseRequirements = requirements
                .Where(t => testCase.RequirementIDs.Any(c => c == t.ID))
                .ToList();

            int Row = 5;
            foreach (var requirement in testCaseRequirements)
            {

                xlsSheet[$"A{Row}"].Value = $"{requirement.ID}";
                xlsSheet[$"C{Row}"].Value = $"{requirement.Objective}";
                xlsSheet[$"U{Row}"].Value = $"{string.Join(", ", requirement.EpicIDs)}";
                Row++;
            }

            foreach (var epicId in testCase.EpicIDs)
            {
                xlsSheet[$"A{Row}"].Value = $"{epicId}";
            }

        }
    }
}
