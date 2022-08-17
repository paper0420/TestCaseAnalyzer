using IronXL;
using SwiftExcel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    internal class FinalReportGenerator
    {

        public static void GenerateReport(SpecParameters spec, string reportType, string carLine)
        {
            Console.WriteLine("**Writing report**");
            int currentRowFunctionSheet = 21;
            int currentRowFuSiSheet = 21;

            var workbook = WorkBook.Load(@"templateReport.xlsm");
            var worksheet = workbook.GetWorkSheet("FuSi");


            foreach (var specTC in spec.TestCases)
            {
                
                if (specTC.Carline.Contains(carLine))
                {
                    var writeCheckTest = false;

                    if (reportType == "Fusa" && (specTC.Type.Contains("L1") || specTC.Type.Contains("L2")))
                    {
                        //Console.WriteLine("Write Testcase ID: " + specTC.ID);

                        writeCheckTest = true;
                    }

                    if (reportType == "HV" && specTC.Type.Contains("L1"))
                    {
                        writeCheckTest = true;
                    }

                    if(reportType == "Full")
                    {
                        writeCheckTest = true;
                    }

                    if (writeCheckTest)
                    {
                        var next = WriteCheckTestCase(spec, ref currentRowFunctionSheet, ref currentRowFuSiSheet, workbook, specTC);

                        if (next)
                        {
                            continue;
                        }
                    }

                }


            }

            workbook.Save();


        }

        private static bool WriteCheckTestCase(
            SpecParameters spec, 
            ref int currentRowFunctionSheet, 
            ref int currentRowFuSiSheet, 
            WorkBook workbook, 
            TestCaseOnlyExecutedItem specTC)
        {
            var htmlTC = spec.HtmlDatasByID.ContainsKey(specTC.ID)
                ? spec.HtmlDatasByID[specTC.ID]
                : null;

            if (htmlTC == null)
            {
                if (specTC.Result != null)
                {
                    WriteRequirements(spec, ref currentRowFunctionSheet, ref currentRowFuSiSheet, specTC, htmlTC, workbook);
                }
                else
                {
                    return true;
                }

            }

            WriteRequirements(spec, ref currentRowFunctionSheet, ref currentRowFuSiSheet, specTC, htmlTC, workbook);
            
            return false;
        }

        private static void WriteRequirements(SpecParameters spec, 
            ref int currentRowFunctionSheet, 
            ref int currentRowFuSiSheet, 
            TestCaseOnlyExecutedItem specTC, 
            HtmlData htmlTC,
            WorkBook workbook)
        {
            foreach (var req in specTC.RequirementIDs)
            {
                var isNumeric = int.TryParse(req, out int n);
                if (isNumeric)
                {
                    var klhID = spec.CurrentRequirementsByID.ContainsKey(req)
                        ? spec.CurrentRequirementsByID[req]
                        : null;

                    if (klhID == null)
                    {
                        continue;
                    }

                    if (klhID.FusaType?.Contains("ASIL", "QM", "n.a.", "SR") == true)
                    {
                        //var workbook = WorkBook.Load(@"tprp.xlsx");
                        //var worksheet = workbook.GetWorkSheet("FuSi");
                        var worksheet = workbook.GetWorkSheet("FuSi");
                        WriteRow(currentRowFuSiSheet, htmlTC, specTC, req, worksheet);
                        //Write(currentRowFuSiSheet, htmlTC, specTC, req, fusiSheet);

                        currentRowFuSiSheet++;
                        //workbook.Save();


                    }
                    else
                    {
                        //var workbook = WorkBook.Load(@"tprp.xlsx");
                        //var worksheet = workbook.GetWorkSheet("Functional");
                        var worksheet = workbook.GetWorkSheet("Functional");
                        WriteRow(currentRowFunctionSheet, htmlTC, specTC, req, worksheet);
                        //Write(currentRowFunctionSheet, htmlTC, specTC, req, functionalSheet);

                        currentRowFunctionSheet++;
                        //workbook.Save();

                    }

                }

            }
        }


        private static void WriteRow(int currentRow, HtmlData htmlTC, TestCaseOnlyExecutedItem specTC, string req, WorkSheet worksheet)
        {
            if(htmlTC != null)
            {
                
                worksheet[$"R{currentRow}"].Value = htmlTC.TotalTestResult;
                worksheet[$"U{currentRow}"].Value = htmlTC.NumberOfPassed;
                worksheet[$"W{currentRow}"].Value = htmlTC.NumberOfFailed;
                worksheet[$"X{currentRow}"].Value = htmlTC.NumberOfNotExecuted;
                if (htmlTC.TotalTestResult == "FAILED")
                {
                    worksheet[$"Q{currentRow}"].Value = specTC.Comment;
                }

            }
            else
            {
                worksheet[$"R{currentRow}"].Value = "No HTML and Reviewed Needed";
            }

            worksheet[$"H{currentRow}"].Value = specTC.ID;
            worksheet[$"B{currentRow}"].Value = specTC.ItemClass1;
            worksheet[$"C{currentRow}"].Value = specTC.ItemClass2;
            worksheet[$"D{currentRow}"].Value = specTC.ItemClass3;
            worksheet[$"I{currentRow}"].Value = specTC.Objective;
            worksheet[$"J{currentRow}"].Value = specTC.Type;

            worksheet[$"K{currentRow}"].Value = req;
            worksheet[$"M{currentRow}"].Value = req;

            
        }



        private static void Write(int currentRow, HtmlData htmlTC, TestCaseOnlyExecutedItem specTC, string req, ExcelWriter sheet)
        {
            if (htmlTC != null)
            {
                sheet.Write(htmlTC.ID, 8, currentRow);
                if (htmlTC.TotalTestResult == "FAILED")
                {
                    sheet.Write(specTC.Comment, 16, currentRow);

                }
                sheet.Write(htmlTC.TotalTestResult, 17, currentRow); 
                sheet.Write(htmlTC.NumberOfPassed.ToString(), 20, currentRow);
                sheet.Write(htmlTC.NumberOfFailed.ToString(), 22, currentRow);
                sheet.Write(htmlTC.NumberOfNotExecuted.ToString(), 23, currentRow);

       

            }
            else
            {
                sheet.Write("No HTML and Reviewed Needed", 17, currentRow);
                
            }

            sheet.Write(specTC.ItemClass1, 2, currentRow);
            sheet.Write(specTC.ItemClass2, 3, currentRow);
            sheet.Write(specTC.ItemClass3, 4, currentRow);
            sheet.Write(specTC.Type, 5, currentRow);
            sheet.Write(specTC.ID, 8, currentRow);
            sheet.Write(specTC.Objective, 9, currentRow);
            sheet.Write(req, 11, currentRow);
            sheet.Write(req, 12, currentRow);

        }

    }
}
