using System.Collections.Generic;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public interface IReportGenerator
    {
        void GenerateReport(List<TestCase> testCases, List<Requirement> requirements);

    }
}
