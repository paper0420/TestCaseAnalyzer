using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class LTestCases
    {
        public static LTestCases ReadL1(IExcelDataReader reader)
        {
            return Read(reader, 2, 10, 11, 12, 14);
        }

        public static LTestCases ReadL2(IExcelDataReader reader)
        {
            return Read(reader, 2, 8, 9, 10, 12);
        }

        public static LTestCases ReadL3(IExcelDataReader reader)
        {
            return Read(reader, 2, 8, 9, 10, 12);
        }

        private static LTestCases Read(
            IExcelDataReader reader,
            int idColumn,
            int testTimeColumn,
            int historyColumn,
            int ticketNumberColumn,
            int commentColumn)
        {
            var l1 = new LTestCases();

            var id = reader.GetValue(idColumn);
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                l1.ID = reader.GetString(idColumn).Trim();

            }

            var testTimeAsString = reader.GetValue(testTimeColumn)?.ToString();
            if (DateTime.TryParse(testTimeAsString, out var testTimeAsDate))
            {
                l1.TestTime = testTimeAsDate.ToString("HH:mm:ss");
            }

            l1.History = reader.GetString(historyColumn);
            l1.TicketNumber = reader.GetString(ticketNumberColumn);
            l1.Comment = reader.GetString(commentColumn);

            return l1;
        }


        public string ID { get; private set; }
        public string TestTime { get; private set; }

        public string History { get; private set; }
        public string TicketNumber { get; private set; }
        public string Comment { get; private set; }
    }
}
