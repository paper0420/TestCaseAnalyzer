using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public class DeletedReq
    {
        public DeletedReq(IExcelDataReader reader)
        {
            var changeStatus = reader.GetString(3);
            //var itemType = reader.GetString(7);
            var id = reader.GetValue(1);
            var panaStatus = reader.GetString(10);


            if (changeStatus == "Deleted")
            {
                if(!string.IsNullOrWhiteSpace(id?.ToString()))
                {
                    this.ID = (int)reader.GetDouble(1);
                    this.Objective = reader.GetString(5);
                    //Console.WriteLine(ID);

                }
                else
                {
                    return;
                }


            }

            
            
        }

             
            
               

     

        
        public int ID { get; }
        public string Objective { get; }
        public string ChangeStatus { get; }
        public string ItemType { get; }

        public string PanaStatus { get; }





    }
}

