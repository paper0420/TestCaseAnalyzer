using ExcelDataReader;
using IronXL;
using System;
using System.Collections.Generic;
using System.IO;

namespace TestCaseAnalyzer.App.FileReader
{
    public class ExcelReader
    {
        public IEnumerable<T> ReadFile<T>(string file,
            string sheet,
            int rowNumber,
            Func<IExcelDataReader, ExcelColumnReader, T> func)
        {

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    do
                    {
                        if (reader.Name == sheet)
                        {
                            if (rowNumber != 1)
                            {
                                for (int i = 0; i < rowNumber; i++)
                                {
                                    reader.Read();

                                }

                            }


                            reader.Read();
                            var index = new ExcelColumnReader(reader, file);

                            while (reader.Read())
                            {
                                var t = func(reader, index);
                                yield return t;
                            }
                        }

                    } while (reader.NextResult());
                }
            }


        }


        



    }
}