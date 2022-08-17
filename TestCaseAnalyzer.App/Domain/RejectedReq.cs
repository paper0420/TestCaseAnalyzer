using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class RejectedReq
    {
        public RejectedReq(IExcelDataReader reader)
        {
            var changeStatus = reader.GetString(3);
            //var itemType = reader.GetString(7);
            var id = reader.GetValue(1);
            var panaStatus = reader.GetString(10);


            if (panaStatus == "rejected")
            {
                if (!string.IsNullOrWhiteSpace(id?.ToString()))
                {
                    this.ID = reader.GetValue(1)?.ToString();
                    this.Objective = reader.GetString(5);
                    //Console.WriteLine(ID);

                }
                else
                {
                    return;
                }


            }



        }
        public string ID { get; }
        public string Objective { get; }
        public string ChangeStatus { get; }
        public string ItemType { get; }
        public string PanaStatus { get; }


    }
}
