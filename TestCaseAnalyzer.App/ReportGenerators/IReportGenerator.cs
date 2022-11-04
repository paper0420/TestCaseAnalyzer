using TestCaseAnalyzer.App.Spec;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public interface IReportGenerator
    {
        void GenerateReport(SpecParameters spec);

    }
}
