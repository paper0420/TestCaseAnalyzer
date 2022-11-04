using ExcelDataReader;
using System;
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

        public static string? GetStringOrNull(this IExcelDataReader reader, int? columnIndex)
        {
            return columnIndex == null
                ? null
                : reader.GetValue(columnIndex.Value)?.ToString();
        }

        public static string SubstringFrom(this string input,
        string value,
        StringComparison stringComparison = StringComparison.OrdinalIgnoreCase,
        bool inclusive = true)
        {
            int index = input.IndexOf(value, stringComparison);

            if (index > -1)
            {
                index = inclusive ? index : index + value.Length;

                return input.Substring(index);
            }

            return string.Empty;
        }

        public static string SubstringUpTo(
            this string input,
            string value,
            StringComparison stringComparison = StringComparison.OrdinalIgnoreCase,
            bool inclusive = false)
        {
            int index = input.IndexOf(value, stringComparison);

            if (index > -1)
            {
                index = inclusive ? index + value.Length : index;

                return input.Substring(0, index);
            }

            return input;
        }
    }
}
