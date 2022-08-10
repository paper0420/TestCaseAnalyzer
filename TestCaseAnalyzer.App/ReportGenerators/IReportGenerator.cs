using System.Collections.Generic;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public interface IReportGenerator
    {
        void GenerateReport(SpecParameters spec);

    }
}
