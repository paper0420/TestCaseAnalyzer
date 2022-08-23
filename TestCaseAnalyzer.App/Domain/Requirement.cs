using ExcelDataReader;
using System;
using System.Collections.Generic;

namespace TestCaseAnalyzer.App
{
    public class Requirement
    {
        public Requirement(IExcelDataReader reader, ExcelColumnReader index)
        {

            this.changeStatus = reader.GetString(index.ChangeStatusIndex);
            this.panaStatus = reader.GetString(index.ChangeStatusIndex);
            this.VerificationSpecStatus = reader.GetValue(index.VerificationSpecStatusIndex)?.ToString();

            this.FusaType = reader.GetValue(index.KLHEAS_ASILIndex)?.ToString();

            var id = reader.GetValue(index.KLHIDIndex);
            if (!string.IsNullOrWhiteSpace(id?.ToString()))
            {
                this.ID = reader.GetValue(index.KLHIDIndex)?.ToString();
            }

            this.Objective = reader.GetString(index.KLHObjectiveIndex);


            this.VerificationMeasure = reader.GetValue(index.VerificationMeasureIndex)?.ToString().Replace("\n", "");
            this.Type = reader.GetValue(4)?.ToString();
            //this.EpicIDs = reader
            //    .GetString(21)?
            //    .Split("\n", System.StringSplitOptions.RemoveEmptyEntries)
            //    ?? new string[0];

        }

        public string ID { get; }
        public string Objective { get; }
        public string changeStatus { get; }
        public string panaStatus { get; }
        public string VerificationSpecStatus { get; }
        public string VerificationMeasure { get; }
        public string Type { get; }
        public string[] EpicIDs { get; }
        public string FusaType { get; }



    }
}