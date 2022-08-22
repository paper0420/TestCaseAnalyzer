using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Domain;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

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
            var fileNameInOutputPaht = Path.GetFullPath(fileName, FileNames.OutputFolder);
            File.Copy(FileNames.ReportTemplateFile, fileNameInOutputPaht);

            //var workbook = WorkBook.Load(fileNameInOutputPaht);

            var openSettings = new OpenSettings()
            {
                RelationshipErrorHandlerFactory = package =>
                {
                    return new UriRelationshipErrorHandler();
                }
            };

            using (var outputStream = File.Open(fileNameInOutputPaht, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (var xls = SpreadsheetDocument.Open(outputStream, true, openSettings))
                {
                    xls.Save();
                }

                var workbook = new XLWorkbook(outputStream);
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

            //foreach (var specTC in spec.TestCases)
            //{
            //    if (!checkTCID.Contains(specTC.ID))
            //    {
            //        if (specTC.Carline.Contains(carLine))
            //        {
            //            var writeCheckTest = false;

            //            if (reportType == "Fusa" && (specTC.Type.Contains("L1") || specTC.Type.Contains("L2")))
            //            {

            //                writeCheckTest = true;
            //            }

            //            if (reportType == "HV" && specTC.Type.Contains("L1"))
            //            {
            //                writeCheckTest = true;
            //            }

            //            if (reportType == "Full")
            //            {
            //                writeCheckTest = true;
            //            }

            //            if (writeCheckTest)
            //            {
            //                var htmlTC = spec.HtmlDatasByID.ContainsKey(specTC.ID)
            //                               ? spec.HtmlDatasByID[specTC.ID]
            //                               : null;

            //                WriteRequirements(spec, ref currentRowFunctionSheet, ref currentRowFuSiSheet, specTC, htmlTC, workbook,
            //                functionalTestCases, fusiTestCases);


            //                checkTCID.Add(specTC.ID);
            //                continue;

            //            }
            //        }
            //    }
            //}
            ////WriteTotalSubTCs(workbook);
            ////WriteTestIdentification(workbook,spec, carLine, swRelease);
            ////WriteTestIden(wb, spec, carLine, swRelease);
            ////wb.Save();
            //workbook.Save();

            //Console.WriteLine("**Report finished : " + now.ToString("F"));
        }


        private static void WriteRequirements(SpecParameters spec,
            ref int currentRowFunctionSheet,
            ref int currentRowFuSiSheet,
            TestCaseOnlyExecutedItem specTC,
            HtmlData htmlTC,
            XLWorkbook workbook,
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
                        //var worksheet = workbook.GetWorkSheet("FuSi");
                        //WriteRow(spec, currentRowFuSiSheet, htmlTC, specTC, klhID, worksheet);
                        var worksheet = workbook.Worksheet("FuSi");

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
                        //var worksheet = workbook.GetWorkSheet("Functional");
                        //WriteRow(spec, currentRowFunctionSheet, htmlTC, specTC, klhID, worksheet);
                        var worksheet = workbook.Worksheet("Functional");
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

        private static void WriteRowOld(SpecParameters spec,
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

        private static void WriteRow(SpecParameters spec,
            int currentRow,
            HtmlData htmlTC,
            TestCaseOnlyExecutedItem specTC,
            Requirement klhID,
            IXLWorksheet worksheet)
        {

            if (htmlTC != null)
            {
                worksheet.Cell($"R{currentRow}").Value = htmlTC.TotalTestResult;
                worksheet.Cell($"U{currentRow}").Value = htmlTC.NumberOfPassed;
                worksheet.Cell($"W{currentRow}").Value = htmlTC.NumberOfFailed;
                worksheet.Cell($"X{currentRow}").Value = htmlTC.NumberOfNotExecuted;

                if (htmlTC.TotalTestResult == "FAILED")
                {
                    worksheet.Cell($"Q{currentRow}").Value = specTC.Comment;
                    //worksheet.Cell($"R{currentRow}").Style.SetBackgroundColor(ColorCodes.Red);
                }

                if (htmlTC.TotalTestResult == "PASSED" && (htmlTC.NumberOfNotExecuted > 0 || htmlTC.NumberOfFailed > 0))
                {
                    worksheet.Cell($"Q{currentRow}").Value = specTC.Comment;
                }

            }

            if (htmlTC == null)
            {
                if (specTC.Result == "OBSOLETE")
                {
                    //worksheet.Cell($"R{currentRow}").Style.SetBackgroundColor(ColorCodes.Grey);
                    worksheet.Cell($"R{currentRow}").Value = specTC.Result;
                    worksheet.Cell($"Q{currentRow}").Value = specTC.Comment;
                }
                else if(specTC.Result.Contains("NOT TESTABLE"))
                {
                    worksheet.Cell($"R{currentRow}").Value = "NOT TESTABLE";
                    worksheet.Cell($"Q{currentRow}").Value = specTC.Comment;
                }
                else
                {
                    worksheet.Cell($"R{currentRow}").Value = "No HTML and Reviewed Needed";
                }

            }

            worksheet.Cell($"A{currentRow}").Value = "MicroFuzzy";
            worksheet.Cell($"B{currentRow}").Value = specTC.ItemClass1;
            worksheet.Cell($"C{currentRow}").Value = specTC.ItemClass2;
            worksheet.Cell($"D{currentRow}").Value = specTC.ItemClass3;
            worksheet.Cell($"H{currentRow}").Value = specTC.ID;


            worksheet.Cell($"I{currentRow}").Value = specTC.Objective;
            worksheet.Cell($"J{currentRow}").Value = specTC.Type;

            worksheet.Cell($"K{currentRow}").Value = klhID.ID;
            worksheet.Cell($"M{currentRow}").Value = klhID.ID;
            worksheet.Cell($"N{currentRow}").Value = "11kW Big";

            worksheet.Cell($"O{currentRow}").Value = specTC.VerificationMethod;
            worksheet.Cell($"P{currentRow}").Value = klhID.VerificationSpecStatus;

            worksheet.Cell($"Z{currentRow}").Value = specTC.TestCatHV;
            worksheet.Cell($"AA{currentRow}").Value = specTC.TestCatBasic;
            worksheet.Cell($"AB{currentRow}").Value = specTC.TestCatFusa;
            worksheet.Cell($"AC{currentRow}").Value = specTC.TestCatFunc;
            worksheet.Cell($"AD{currentRow}").Value = specTC.TestCatFull;



            if (worksheet.Name == "FuSi")
            {
                foreach (var sg in spec.SafetyGoalKLHs)
                {
                    if (sg.ID == klhID.ID)
                    {
                        worksheet.Cell($"L{currentRow}").Value = sg.SG;
                        break;
                    }
                }

            }


        }

        private static void WriteTotalSubTCs(XLWorkbook workbook)
        {
            var worksheetFusi = workbook.Worksheet("FuSi");
            var worksheetFunc = workbook.Worksheet("Functional");

            worksheetFusi.Cell($"S5").Value = $"Not Executed: {TotalSubTestCases.fusiTotalSubTestcaseNotExecuted}";
            worksheetFusi.Cell($"S6").Value = $"PASSED: {TotalSubTestCases.fusiTotalSubTestcasePassed}";
            worksheetFusi.Cell($"S8").Value = $"FAILED: {TotalSubTestCases.fusiTotalSubTestcaseFailed}";

            worksheetFunc.Cell($"S5").Value = $"Not Executed: {TotalSubTestCases.funcTotalSubTestcaseNotExecuted}";
            worksheetFunc.Cell($"S6").Value = $"PASSED: {TotalSubTestCases.funcTotalSubTestcasePassed}";
            worksheetFunc.Cell($"S8").Value = $"FAILED: {TotalSubTestCases.funcTotalSubTestcaseFailed}";
        }

        private static void WriteTestIdentificationOld(WorkBook workbook, SpecParameters spec, string carLine, string swRelease)
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

        private static void WriteTestIdentification(XLWorkbook workbook, SpecParameters spec, string carLine, string swRelease)
        {
            var htmlTC = spec.HtmlDatas.First();

            var worksheet = workbook.Worksheet("Test_Identification");
            worksheet.Cell("B19").Value = htmlTC.HW;
            worksheet.Cell("B23").Value = $"SW {carLine} {swRelease}";
            worksheet.Cell("B25").Value = htmlTC.BTLD;
            worksheet.Cell("B25").Value = htmlTC.BTLD;
            worksheet.Cell("B26").Value = htmlTC.SWFL;
            worksheet.Cell("B29").Value = htmlTC.PIC;
            worksheet.Cell("B30").Value = htmlTC.DSP;
        }




    }
}
