using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.FileReader
{
    public class ExcelReader
    {
        public WorksheetData<T> ReadFile<T>(string file,
            string sheet,
            Func<IExcelDataReader, Header, T> func)
        {
            //Console.WriteLine(AppDomain.CurrentDomain.BaseDirectory);
            //file = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, file));
            //Console.WriteLine(file);

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                var worksheet = new WorksheetData<T>();
                
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        if (reader.Name == sheet || reader.Name.Contains(sheet))
                        {
                            reader.Read();
                            
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                var column = new Column
                                {
                                    Index = i,
                                    Name = reader.GetValue(i)?.ToString() 
                                };
                                
                                worksheet.Header.Columns.Add(column);

                            }

                            if (sheet == "Test_Item")
                            {
                                worksheet.Header.GetCarLineNames(worksheet.Header.Columns);

                            }



                            while (reader.Read())
                            {
                                var row = func(reader, worksheet.Header);
                                worksheet.DataRows.Add(row);
                            }
                        }

                    } while (reader.NextResult());
                }

                return worksheet;
            }
        }


        public IEnumerable<T> ReadPanaFile<T>(string file,
            string sheet,
            int rowNumber,
            Func<IExcelDataReader, T> func)
        {

            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    do
                    {
                        if (reader.Name == sheet || reader.Name.Contains(sheet))
                        {
                            if (rowNumber != 1)
                            {
                                for (int i = 1; i < rowNumber; i++)
                                {
                                    reader.Read();

                                }

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