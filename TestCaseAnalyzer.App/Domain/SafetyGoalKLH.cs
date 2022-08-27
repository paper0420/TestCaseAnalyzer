using ExcelDataReader;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public class SafetyGoalKLH
    {
        public SafetyGoalKLH(IExcelDataReader reader, Header header)
        {
            SG = reader.GetValue(header.GetColumnIndex("SG"))?.ToString();
            ID = reader.GetValue(header.GetColumnIndex("Object ID from Original"))?.ToString();
        }

        public string SG { get; }
        public string ID { get; }
        public string Attribute { get; }
        public string Asil { get; }
    }
}