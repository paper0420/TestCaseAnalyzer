using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Domain;
using System.IO;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    internal class FinalReportGenerator
    {

        public static void GenerateReport(SpecParameters spec, string reportType, string carLine, string swRelease)
        {
            DateTime now = DateTime.Now;
            Console.WriteLine("***Writing report***");

            int currentRowFunctionSheet = 21;
            int currentRowFuSiSheet = 21;

            var functionalTestCases = new HashSet<string>();
            var fusiTestCases = new HashSet<string>();
            TotalSubTestCases.funcTotalSubTestcasePassed = 0;
            TotalSubTestCases.funcTotalSubTestcaseFailed = 0;
            TotalSubTestCases.funcTotalSubTestcaseNotExecuted = 0;
            TotalSubTestCases.fusiTotalSubTestcasePassed = 0;
            TotalSubTestCases.fusiTotalSubTestcaseFailed = 0;
            TotalSubTestCases.fusiTotalSubTestcaseNotExecuted = 0;
            var checkTCID = new HashSet<string>();

            var reportTypeforFileName = reportType;
            if (reportType == "Fusa")
            {
                reportTypeforFileName = "FusaBasic";
            }
            if (reportType == "HV")
            {
                reportTypeforFileName = "HVRTU";
            }

            var fileName = $"BMW_CCU_SystemTestReport_SW{swRelease}_{carLine}_{reportTypeforFileName}_{now.ToString("ddHHmmss")}.xlsx";
            var fileNameInOutputPaht = $"..//..//..//Output//{fileName}";
            File.Copy(FileNames.ReportTemplateFile, fileNameInOutputPaht);

            var workbook = WorkBook.Load(fileNameInOutputPaht);

            foreach (var specTC in spec.TestCases)
            {
                if (!checkTCID.Contains(specTC.ID))
                {
                    if (specTC.Carline.Contains(carLine))
                    {
                        var writeCheckTest = false;

                        if (reportType == "Fusa" && (specTC.Type.Contains("L1") || specTC.Type.Contains("L2")))
                        {

                            writeCheckTest = true;
                        }

                        if (reportType == "HV" && specTC.Type.Contains("L1"))
                        {
                            writeCheckTest = true;
                        }

                        if (reportType == "Full")
                        {
                            writeCheckTest = true;
                        }

                        if (writeCheckTest)
                        {
                            var htmlTC = spec.HtmlDatasByID.ContainsKey(specTC.ID)
                                           ? spec.HtmlDatasByID[specTC.ID]
                                           : null;

                            WriteRequirements(spec, ref currentRowFunctionSheet, ref currentRowFuSiSheet, specTC, htmlTC, workbook,
                            functionalTestCases, fusiTestCases);


                            checkTCID.Add(specTC.ID);
                            continue;

                        }
                    }
                }
            }
            WriteTotalSubTCs(workbook);
            WriteTestIdentification(workbook,spec, carLine, swRelease);
            workbook.Save();
            Console.WriteLine("**Report finished : " + now.ToString("F"));
        }

        private static void WriteRequirements(SpecParameters spec,
            ref int currentRowFunctionSheet,
            ref int currentRowFuSiSheet,
            TestCaseOnlyExecutedItem specTC,
            HtmlData htmlTC,
            WorkBook workbook,
            HashSet<string> functionalTestCases,
            HashSet<string> fusiTestCases)
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
                        var worksheet = workbook.GetWorkSheet("FuSi");
                        WriteRow(spec, currentRowFuSiSheet, htmlTC, specTC, klhID, worksheet);
                        currentRowFuSiSheet++;

                        if (!fusiTestCases.Contains(specTC.ID))
                        {
                            CalculateTotalSubTCsforFusiSheet(specTC, htmlTC);
                        }

                        fusiTestCases.Add(specTC.ID);
                    }
                    else
                    {
                        var worksheet = workbook.GetWorkSheet("Functional");
                        WriteRow(spec, currentRowFunctionSheet, htmlTC, specTC, klhID, worksheet);
                        currentRowFunctionSheet++;

                        if (!functionalTestCases.Contains(specTC.ID))
                        {
                            CalculateTotalSubTCsforFuncSheet(specTC, htmlTC);
                        }

                        functionalTestCases.Add(specTC.ID);
                    }
                }
            }
            
        }

        private static void CalculateTotalSubTCsforFusiSheet(TestCaseOnlyExecutedItem specTC, HtmlData htmlTC)
        {
            if (htmlTC != null)
            {
                TotalSubTestCases.fusiTotalSubTestcasePassed += htmlTC.NumberOfPassed;
                TotalSubTestCases.fusiTotalSubTestcaseFailed += htmlTC.NumberOfFailed;
                TotalSubTestCases.fusiTotalSubTestcaseNotExecuted += htmlTC.NumberOfNotExecuted;

            }

        }

        private static void CalculateTotalSubTCsforFuncSheet(TestCaseOnlyExecutedItem specTC, HtmlData htmlTC)
        {
            if (htmlTC != null)
            {
                TotalSubTestCases.funcTotalSubTestcasePassed += htmlTC.NumberOfPassed;
                TotalSubTestCases.funcTotalSubTestcaseFailed += htmlTC.NumberOfFailed;
                TotalSubTestCases.funcTotalSubTestcaseNotExecuted += htmlTC.NumberOfNotExecuted;

            }

        }

        private static void WriteRow(SpecParameters spec,
            int currentRow,
            HtmlData htmlTC,
            TestCaseOnlyExecutedItem specTC,
            Requirement klhID,
            WorkSheet worksheet)
        {
            if (htmlTC != null)
            {
                worksheet[$"R{currentRow}"].Value = htmlTC.TotalTestResult;
                worksheet[$"U{currentRow}"].Value = htmlTC.NumberOfPassed;
                worksheet[$"W{currentRow}"].Value = htmlTC.NumberOfFailed;
                worksheet[$"X{currentRow}"].Value = htmlTC.NumberOfNotExecuted;

                if (htmlTC.TotalTestResult == "FAILED")
                {
                    worksheet[$"Q{currentRow}"].Value = specTC.Comment;
                    worksheet[$"R{currentRow}"].Style.SetBackgroundColor(ColorCodes.Red);
                }

                if (htmlTC.TotalTestResult == "PASSED" && (htmlTC.NumberOfNotExecuted > 0 || htmlTC.NumberOfFailed > 0))
                {
                    worksheet[$"Q{currentRow}"].Value = specTC.Comment;
                }

            }

            if (htmlTC == null)
            {
                if (specTC.Result == "OBSOLETE")
                {
                    worksheet[$"R{currentRow}"].Style.SetBackgroundColor(ColorCodes.Grey);
                    worksheet[$"R{currentRow}"].Value = specTC.Result;
                }
                else
                {
                    worksheet[$"R{currentRow}"].Value = "No HTML and Reviewed Needed";
                }

            }

            worksheet[$"A{currentRow}"].Value = "MicroFuzzy";
            worksheet[$"B{currentRow}"].Value = specTC.ItemClass1;
            worksheet[$"C{currentRow}"].Value = specTC.ItemClass2;
            worksheet[$"D{currentRow}"].Value = specTC.ItemClass3;
            worksheet[$"H{currentRow}"].Value = specTC.ID;


            worksheet[$"I{currentRow}"].Value = specTC.Objective;
            worksheet[$"J{currentRow}"].Value = specTC.Type;

            worksheet[$"K{currentRow}"].Value = klhID.ID;
            worksheet[$"M{currentRow}"].Value = klhID.ID;
            worksheet[$"N{currentRow}"].Value = "11kW Big";

            worksheet[$"O{currentRow}"].Value = specTC.VerificationMethod;
            worksheet[$"P{currentRow}"].Value = klhID.VerificationSpecStatus;

            worksheet[$"Z{currentRow}"].Value = specTC.TestCatHV;
            worksheet[$"AA{currentRow}"].Value = specTC.TestCatBasic;
            worksheet[$"AB{currentRow}"].Value = specTC.TestCatFusa;
            worksheet[$"AC{currentRow}"].Value = specTC.TestCatFunc;
            worksheet[$"AD{currentRow}"].Value = specTC.TestCatFull;



            if (worksheet.Name == "FuSi")
            {
                foreach (var sg in spec.SafetyGoalKLHs)
                {
                    if (sg.ID == klhID.ID)
                    {
                        worksheet[$"L{currentRow}"].Value = sg.SG;
                        break;
                    }
                }

            }


        }

        private static void WriteTotalSubTCs(WorkBook workbook)
        {
            var worksheetFusi = workbook.GetWorkSheet("FuSi");
            var worksheetFunc = workbook.GetWorkSheet("Functional");

            worksheetFusi[$"S5"].Value = $"Not Executed: {TotalSubTestCases.fusiTotalSubTestcaseNotExecuted}";
            worksheetFusi[$"S6"].Value = $"PASSED: {TotalSubTestCases.fusiTotalSubTestcasePassed}";
            worksheetFusi[$"S8"].Value = $"FAILED: {TotalSubTestCases.fusiTotalSubTestcaseFailed}";

            worksheetFunc[$"S5"].Value = $"Not Executed: {TotalSubTestCases.funcTotalSubTestcaseNotExecuted}";
            worksheetFunc[$"S6"].Value = $"PASSED: {TotalSubTestCases.funcTotalSubTestcasePassed}";
            worksheetFunc[$"S8"].Value = $"FAILED: {TotalSubTestCases.funcTotalSubTestcaseFailed}";
        }

        private static void WriteTestIdentification(WorkBook workbook,SpecParameters spec,string carLine , string swRelease)
        {
            var htmlTC = spec.HtmlDatas.First();
            var worksheet = workbook.GetWorkSheet("Test_Identification");
            worksheet["B19"].Value = htmlTC.HW;
            worksheet["B23"].Value = $"SW {carLine} {swRelease}";
            worksheet["B25"].Value = htmlTC.BTLD;
            worksheet["B26"].Value = htmlTC.SWFL;
            worksheet["B29"].Value = htmlTC.PIC;
            worksheet["B30"].Value = htmlTC.DSP;

        }



    }
}
