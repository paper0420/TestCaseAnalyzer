using IronXL;
using System.Collections.Generic;
using System.Linq;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public class TestcasesAndRequirementExcelGenerator
    {
        public static void CreateTestCaseExcel(
            List<Requirement> requirements,
           TestCaseOnlyExecutedItem testCase,
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
            xlsSheet["D2"].Value = $"[{testCase.ItemClass1}],[{testCase.ItemClass2}],[{testCase.ItemClass3}],[{testCase.Type}]";
            xlsSheet["A3"].Value = "Requirements";
        }

        private static void GenerateRequirementsExcel(List<Requirement> requirements, TestCaseOnlyExecutedItem testCase, WorkSheet xlsSheet)
        {
            var testCaseRequirements = requirements
                .Where(requirement => testCase.RequirementIDs.Any(tcReqID => tcReqID == requirement.ID))
                .ToList();

            
            int row = 5;
            foreach (var requirement in testCaseRequirements)
            {

                //    xlsSheet[$"A{Row}"].Value = $"{testCase.ID}";
                //    xlsSheet[$"B{Row}"].Value = $"{testCase.Objective}";
                xlsSheet[$"C{row}"].Value = $"{requirement.ID},[{requirement.changeStatus}],[{requirement.panaStatus}]";
                xlsSheet[$"D{row}"].Value = $"{requirement.Objective}";
                //xlsSheet[$"E{Row}"].Value = $"{string.Join(", ", requirement.EpicIDs)}";
                row++;
            }


        }
    }
}
