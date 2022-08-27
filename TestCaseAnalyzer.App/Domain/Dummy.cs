using ExcelDataReader;

namespace TestCaseAnalyzer.App
{
    public class Dummy
    {
        public Dummy(IExcelDataReader reader)
        {
            //this.SYR_ID = reader.GetValue(7)?.ToString();
            this.KLH_ID = reader.GetValue(1)?.ToString();

        }

        public string SYR_ID { get; }
        public string KLH_ID { get; }


    }
}
