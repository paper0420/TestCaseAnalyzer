using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App
{
    public class DataFileReader
    {
        public IEnumerable<T> ReadFile<T>(string file, string sheet, int rowNumber, Func<IExcelDataReader, T> func)
        {
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    do
                    {
                        if (reader.Name == sheet)
                        {

                            for (int i = 1; i < rowNumber; i++)
                            {
                                reader.Read();
                            }



                            while (reader.Read())
                            {
                                var t = func(reader);
                                yield return t;
                            }
                        }

                    } while (reader.NextResult());
                }
            }
        }

    }
}