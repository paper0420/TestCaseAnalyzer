using ExcelDataReader;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public static class Extensions
    {
        public static int ToInt(this string value)
        {
            return int.Parse(value);
        }

        public static string Format(this string value, string format)
        {
            return string.Format(format, value);
        }

        public static bool Contains(this string value, params string[] values)
        {
            if (values == null)
            {
                return false;
            }

            foreach (var item in values)
            {
                if (value?.Contains(item) == true)
                {
                    return true;
                }
            }

            return false;
        }

        public static string GetValueAsString(this IExcelDataReader reader, Header header, string columnName)
        {
            return reader.GetValue(header.GetColumnIndex(columnName))?.ToString();
        }
    }
}
