using IronXL;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    internal class FinalReportGenerator
    {
        public static void GenerateFinalReport(SpecParameters baseSpec)
        {
            
           
            //WriteTestCases(baseSpec,"L3", "Functional");
            WriteTestCases(baseSpec, "L1", "FuSi");
            //WriteTestCases(baseSpec, "L2", "FuSi");

        }

        private static void WriteTestCases(SpecParameters baseSpec, string Type, string workSheetName)
        {
            var workbook = WorkBook.Load(@"SystemTestReport.xlsm");
            var worksheet = workbook.GetWorkSheet("FuSi");
            int firstRow = 21;
            foreach (var htmlTC in baseSpec.HtmlDatas)
            {

                foreach (var specTC in baseSpec.TestCases)
                {
                    if (specTC.ID == htmlTC.ID)
                    {
                        //if(specTC.Type == Type)
                        {
                            foreach (var req in specTC.RequirementIDs)
                            {
                                worksheet[$"H{firstRow}"].Value = htmlTC.ID;
                                worksheet[$"R{firstRow}"].Value = htmlTC.TotalTestResult;
                                worksheet[$"U{firstRow}"].Value = htmlTC.NumberOfPassed;
                                worksheet[$"W{firstRow}"].Value = htmlTC.NumberOfFailed;
                                worksheet[$"X{firstRow}"].Value = htmlTC.NumberOfNotExecuted;

                                worksheet[$"B{firstRow}"].Value = specTC.ItemClass1;
                                worksheet[$"C{firstRow}"].Value = specTC.ItemClass2;
                                worksheet[$"D{firstRow}"].Value = specTC.ItemClass3;
                                worksheet[$"I{firstRow}"].Value = specTC.Objective;
                                worksheet[$"J{firstRow}"].Value = specTC.Type;

                                worksheet[$"K{firstRow}"].Value = req;
                                worksheet[$"M{firstRow}"].Value = req;

                                firstRow++;
                            }

                        }
               

                    }
                }
            }
            workbook.Save();
        }
    }
}
