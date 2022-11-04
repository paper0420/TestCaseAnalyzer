using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestCaseAnalyzer.App.Domain;
using TestCaseAnalyzer.App.Spec;
using Path = System.IO.Path;

namespace TestCaseAnalyzer.App.ReportGenerators
{
    public static class ENG9FinalReportGenerator
    {
        public static void ENG9GenerateReport(ENG9Spec spec, string reportType, string carLine, string swRelease)
        {
            DateTime now = DateTime.Now;
            Console.WriteLine("***Writing report***");

            int currentRowFunctionSheet = 4;
            int currentRowFuSiSheet = 4;
            int currentRowFTTSheet = 4;

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
            if (reportType == "FTT")
            {
                reportTypeforFileName = "FTT";
            }


            var fileName = $"BMW_CCU_Report_SW{swRelease}_{carLine}_{reportTypeforFileName}_{now.ToString("ddHHmmss")}.xlsx";
            var fileNameInOutputPaht = Path.GetFullPath(fileName, FileNames.OutputFolder);
            File.Copy(FileNames.ENG9ReportTemplateFile, fileNameInOutputPaht);


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
                        if(specTC.Carlines != null)
                        {
                            if (specTC.Carlines.Any(t => t.Contains(carLine)))
                            {
                                var writeCheckTest = false;

                                if (reportType == "Fusa" && specTC.Catagory.Contains("Fusa"))
                                {
                                    writeCheckTest = true;
                                }

                                if (reportType == "FTT" && specTC.Catagory.Contains("FTT"))
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

                                    if (specTC.Catagory == "Fusa")
                                    {
                                        var worksheet = workbook.Worksheet("Outline of FuSa Test Cases");
                                        WriteFusaSheet(worksheet, spec, specTC, ref currentRowFuSiSheet,htmlTC);

                                    }

                                    if (specTC.Catagory == "Functional")
                                    {
                                        var worksheet = workbook.Worksheet("Outline of Functional Test Case");
                                        WriteFunctionSheet(worksheet, spec, specTC, ref currentRowFunctionSheet, htmlTC);

                                    }

                                    if (specTC.Catagory == "FTT")
                                    {
                                        var worksheet = workbook.Worksheet("Outline of Fault Injection");
                                        WriteFTTSheet(worksheet, spec, specTC, ref currentRowFTTSheet, htmlTC);

                                    }

                                    checkTCID.Add(specTC.ID);
                                    continue;

                                }
                            }

                        }
                        
                    }
                }
                //WriteTotalSubTCs(workbook);
                //WriteTestIdentification(workbook, spec, carLine, swRelease);
                workbook.Save();
                Console.WriteLine("**Report finished : " + now.ToString("F"));
            }
        }

        private static void WriteFusaSheet(IXLWorksheet worksheet,
            ENG9Spec spec,
            ENG9Testcase specTC, 
            ref int currentRowFuSiSheet,
            HtmlData htmlIData)
        {
            if(specTC.TSRID == null)
            {
                worksheet.Cell($"E{currentRowFuSiSheet}").Value = specTC.ID;
                worksheet.Cell($"G{currentRowFuSiSheet}").Value = String.Join("\n", specTC.KLHID);

                if (htmlIData != null)
                {
                    worksheet.Cell($"H{currentRowFuSiSheet}").Value = htmlIData.NumberOfexecuted;
                    worksheet.Cell($"J{currentRowFuSiSheet}").Value = htmlIData.NumberOfexecuted;
                    worksheet.Cell($"K{currentRowFuSiSheet}").Value = htmlIData.TotalTestResult;
                    worksheet.Cell($"L{currentRowFuSiSheet}").Value = htmlIData.NumberOfPassed;
                    worksheet.Cell($"M{currentRowFuSiSheet}").Value = htmlIData.NumberOfFailed;

                }

                if (specTC.ICSID != null)
                {
                    worksheet.Cell($"D{currentRowFuSiSheet}").Value = String.Join("\n", specTC.ICSID);
                }

                currentRowFuSiSheet++;
            }
            else
            {
                foreach (var req in specTC.TSRID)
                {
                    worksheet.Cell($"E{currentRowFuSiSheet}").Value = specTC.ID;
                    worksheet.Cell($"C{currentRowFuSiSheet}").Value = req;
                    worksheet.Cell($"A{currentRowFuSiSheet}").Value = specTC.SafetyGoal;
                    worksheet.Cell($"B{currentRowFuSiSheet}").Value = specTC.RelatedTSRID;
                    worksheet.Cell($"F{currentRowFuSiSheet}").Value = specTC.SubIDs != null ? String.Join("\n", specTC.SubIDs) : null;
                    worksheet.Cell($"G{currentRowFuSiSheet}").Value = specTC.KLHID != null ? String.Join("\n", specTC.KLHID) : null;



                    if (htmlIData != null)
                    {
                        worksheet.Cell($"H{currentRowFuSiSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"J{currentRowFuSiSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"K{currentRowFuSiSheet}").Value = htmlIData.TotalTestResult;
                        worksheet.Cell($"L{currentRowFuSiSheet}").Value = htmlIData.NumberOfPassed;
                        worksheet.Cell($"M{currentRowFuSiSheet}").Value = htmlIData.NumberOfFailed;

                    }



                    if (specTC.ICSID != null)
                    {
                        worksheet.Cell($"D{currentRowFuSiSheet}").Value = String.Join("\n", specTC.ICSID);
                    }
                    currentRowFuSiSheet++;
                }
            }
            

        }

        private static void WriteFunctionSheet(IXLWorksheet worksheet, 
            ENG9Spec spec,
            ENG9Testcase specTC,
            ref int currentRowFunctionSheet,
            HtmlData htmlIData)
        {
            if (specTC.SYRID == null)
            {
                if(specTC.KLHID != null)
                {
                    worksheet.Cell($"C{currentRowFunctionSheet}").Value = specTC.ID;
                    worksheet.Cell($"B{currentRowFunctionSheet}").Value = String.Join("\n", specTC.KLHID);

                    if (htmlIData != null)
                    {
                        worksheet.Cell($"E{currentRowFunctionSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"F{currentRowFunctionSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"G{currentRowFunctionSheet}").Value = htmlIData.TotalTestResult;
                        worksheet.Cell($"H{currentRowFunctionSheet}").Value = htmlIData.NumberOfPassed;
                        worksheet.Cell($"I{currentRowFunctionSheet}").Value = htmlIData.NumberOfFailed;

                    }
                    currentRowFunctionSheet++;
                }

            }
            else
            {
                foreach (var req in specTC.SYRID)
                {
                    worksheet.Cell($"C{currentRowFunctionSheet}").Value = specTC.ID;
                    worksheet.Cell($"A{currentRowFunctionSheet}").Value = req;
                    worksheet.Cell($"D{currentRowFunctionSheet}").Value = specTC.FunctionCatagory;

                    if (htmlIData != null)
                    {
                        worksheet.Cell($"E{currentRowFunctionSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"F{currentRowFunctionSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"G{currentRowFunctionSheet}").Value = htmlIData.TotalTestResult;
                        worksheet.Cell($"H{currentRowFunctionSheet}").Value = htmlIData.NumberOfPassed;
                        worksheet.Cell($"I{currentRowFunctionSheet}").Value = htmlIData.NumberOfFailed;

                    }
                    currentRowFunctionSheet++;

                }

            }
  

        }
        private static void WriteFTTSheet(IXLWorksheet worksheet, 
            ENG9Spec spec, 
            ENG9Testcase specTC, 
            ref int currentRowFTTSheet,
            HtmlData htmlIData)
        {
            if(specTC.ICSID == null)
            {
                worksheet.Cell($"H{currentRowFTTSheet}").Value = specTC.ID;
                if (htmlIData != null)
                {
                    worksheet.Cell($"J{currentRowFTTSheet}").Value = htmlIData.NumberOfexecuted;
                    worksheet.Cell($"K{currentRowFTTSheet}").Value = htmlIData.NumberOfexecuted;
                    worksheet.Cell($"L{currentRowFTTSheet}").Value = htmlIData.TotalTestResult;
                    worksheet.Cell($"M{currentRowFTTSheet}").Value = htmlIData.NumberOfPassed;
                    worksheet.Cell($"N{currentRowFTTSheet}").Value = htmlIData.NumberOfFailed;

                }


                currentRowFTTSheet++;

            }
            else
            {
                foreach (var req in specTC.ICSID)
                {
                    worksheet.Cell($"H{currentRowFTTSheet}").Value = specTC.ID;
                    worksheet.Cell($"C{currentRowFTTSheet}").Value = req;
                    worksheet.Cell($"E{currentRowFTTSheet}").Value = specTC.TSRID;
                    worksheet.Cell($"F{currentRowFTTSheet}").Value = specTC.RelatedTSRID;
                    worksheet.Cell($"G{currentRowFTTSheet}").Value = specTC.ErrorFactor;

                    if (htmlIData != null)
                    {
                        worksheet.Cell($"J{currentRowFTTSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"K{currentRowFTTSheet}").Value = htmlIData.NumberOfexecuted;
                        worksheet.Cell($"L{currentRowFTTSheet}").Value = htmlIData.TotalTestResult;
                        worksheet.Cell($"M{currentRowFTTSheet}").Value = htmlIData.NumberOfPassed;
                        worksheet.Cell($"N{currentRowFTTSheet}").Value = htmlIData.NumberOfFailed;
                    }

                    currentRowFTTSheet++;
                }

            }
       

        }

        //private static void WriteRequirements(ENG9Spec spec,
        //    ref int currentRowFunctionSheet,
        //    ref int currentRowFuSiSheet,
        //    ref int currentRowFTTSheet,
        //    ENG9Testcase specTC,
        //    HtmlData htmlTC,
        //    XLWorkbook workbook,
        //    HashSet<string> functionalTestCases,
        //    HashSet<string> fusiTestCases)
        //{

        //    foreach (var req in specTC.RequirementIDs)
        //    {
        //        var isNumeric = int.TryParse(req, out int n);
        //        if (isNumeric)
        //        {
        //            var klhID = spec.CurrentRequirementsByID.ContainsKey(req)
        //                ? spec.CurrentRequirementsByID[req]
        //                : null;

        //            if (klhID == null)
        //            {
        //                continue;
        //            }

        //            if (klhID.FusaType?.Contains("ASIL", "QM", "n.a.", "SR") == true)
        //            {

        //                var worksheet = workbook.Worksheet("FuSi");
        //                WriteRow(spec, currentRowFuSiSheet, htmlTC, specTC, klhID, worksheet);
        //                currentRowFuSiSheet++;

        //                if (!fusiTestCases.Contains(specTC.ID))
        //                {
        //                    CalculateTotalSubTCsforFusiSheet(specTC, htmlTC);
        //                }

        //                fusiTestCases.Add(specTC.ID);
        //            }
        //            else
        //            {
        //                var worksheet = workbook.Worksheet("Functional");
        //                WriteRow(spec, currentRowFunctionSheet, htmlTC, specTC, klhID, worksheet);

        //                currentRowFunctionSheet++;

        //                if (!functionalTestCases.Contains(specTC.ID))
        //                {
        //                    CalculateTotalSubTCsforFuncSheet(specTC, htmlTC);
        //                }

        //                functionalTestCases.Add(specTC.ID);
        //            }
        //        }
        //    }

        //}

        //private static void WriteRow(IXLWorksheet worksheet, ENG9Testcase specTC,int currentRow)
        //{
        //    worksheet.Cell($"R{currentRow}").Value = "JUSTIFIED";
        //    worksheet.Cell($"Q{currentRow}").Value = specTC.Comment;
        //}
    }
}
