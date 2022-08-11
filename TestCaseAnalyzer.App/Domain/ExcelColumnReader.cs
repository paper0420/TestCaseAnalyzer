using ExcelDataReader;
using System;

namespace TestCaseAnalyzer.App
{
    public static class ExcelColumnReader
    {
        public static int GetColumnIndex(string columnName, IExcelDataReader reader)
        {
            var name = "";
            for(var i = 0; i <= 20; i++)
            {
                name = reader.GetValue(i)?.ToString();
                Console.WriteLine(name);
                if(name == columnName)
                {
                    Console.WriteLine(i);
                    return i;
                }
            }
            return 100;
        }
    }
}