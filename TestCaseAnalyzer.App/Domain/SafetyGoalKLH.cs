using ExcelDataReader;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public class SafetyGoalKLH
    {
        public static SafetyGoalKLH? CreateOrNull(IExcelDataReader reader, Header header)
        {
            var id = reader.GetValue(header.GetColumnIndex("Object ID from Original"))?.ToString();

            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            var safety = new SafetyGoalKLH();

            safety.ID = id;
            safety.SG = reader.GetValue(header.GetColumnIndex("SG"))?.ToString();

            return safety;
        }

        public string? SG { get; private set; }
        public string ID { get; private set; }
        public string? Attribute { get; private set; }
        public string? Asil { get; private set; }
    }
}