using ExcelDataReader;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App
{
    public class DataFileReader
    {
        public IEnumerable<Requirement> ReadRequirements()
        {
            using (var stream = File.Open("data.xlsm", FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    do
                    {
                        if (reader.Name == "KLH_BL10.4")
                        {
                            reader.Read();

                            while (reader.Read())
                            {
                                var t = new Requirement(reader);
                                yield return t;
                            }
                        }

                    } while (reader.NextResult());
                }
            }
        }

        public IEnumerable<TestCase> ReadTestCases()
        {
            using (var stream = File.Open("data.xlsm", FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        if (reader.Name == "5.テスト項目Test item")
                        {
                            for (int i = 1; i < 16; i++)
                            {
                                reader.Read();
                            }

                            while (reader.Read())
                            {
                                var t = new TestCase(reader);
                                yield return t;
                            }
                        }

                    } while (reader.NextResult());

                }
            }
        }
    }
}