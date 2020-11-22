using System.Linq;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var reader = new DataFileReader();
            var testCases = reader.ReadTestCases().ToList();
            var requirements = reader.ReadRequirements().ToList();

            new HtmlReportGenerator().GenerateReport(testCases, requirements);
            new ExcelsReportGenerator().GenerateReport(testCases, requirements);
            new ExcelReportWithAWorkbookGenerator().GenerateReport(testCases, requirements);
       
        }


    }
}
