using System;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.ReportGenerators;

namespace TestCaseAnalyzer.App
{
    class Program
    {
        static void Main(string[] args)
        {
            App.RunMyApp(
                FileNames.KlhFile,
                "KLH10_14.xlsx");
           
            //var g08LCITestcasesL1 = reader.ReadFile("SW.2041.121_G08LCI_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var g08LCITestcasesL2 = reader.ReadFile("SW.2041.121_G08LCI_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var g08LCITestcasesL3 = reader.ReadFile("SW.2041.121_G08LCI_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();

            //var g26TestcasesL1 = reader.ReadFile("SW.2041.200_G26_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var g26TestcasesL2 = reader.ReadFile("SW.2041.200_G26_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var g26TestcasesL3 = reader.ReadFile("SW.2041.200_G26_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();

            //var g28TestcasesL1 = reader.ReadFile("SW.2041.120_G28_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var g28TestcasesL2 = reader.ReadFile("SW.2041.120_G28_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var g28TestcasesL3 = reader.ReadFile("SW.2041.120_G28_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();

            //var g60TestcasesL1 = reader.ReadFile("SW.2149.0_G60_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var g60TestcasesL2 = reader.ReadFile("SW.2149.0_G60_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var g60TestcasesL3 = reader.ReadFile("SW.2149.0_G60_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();

            //var g70TestcasesL1 = reader.ReadFile("SW.2137.21_G70_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var g70TestcasesL2 = reader.ReadFile("SW.2137.21_G70_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var g70TestcasesL3 = reader.ReadFile("SW.2137.21_G70_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();

            //var i20TestcasesL1 = reader.ReadFile("SW.2041.121_I20_Top_Tracking.xlsm", "HVRTU", 4, t => LTestCases.ReadL1(t)).ToList();
            //var i20TestcasesL2 = reader.ReadFile("SW.2041.121_I20_Top_Tracking.xlsm", "FusaBasic", 4, t => LTestCases.ReadL2(t)).ToList();
            //var i20TestcasesL3 = reader.ReadFile("SW.2041.121_I20_Top_Tracking.xlsm", "FullTest", 4, t => LTestCases.ReadL3(t)).ToList();




            //var allCarlineTestcaseDetails = new[]
            //{
            //    g08LCITestcasesL1, //0
            //    g26TestcasesL1,//1
            //    g28TestcasesL1,//2
            //    g60TestcasesL1,//3
            //    g70TestcasesL1,//4
            //    i20TestcasesL1,//5

            //    g08LCITestcasesL2, //6
            //    g26TestcasesL2,//7
            //    g28TestcasesL2,//8
            //    g60TestcasesL2,//9
            //    g70TestcasesL2,//10
            //    i20TestcasesL2,//11

            //    g08LCITestcasesL3,//12
            //    g26TestcasesL3,//13               
            //    g28TestcasesL3,//14  
            //    g60TestcasesL3,//15
            //    g70TestcasesL3,//16
            //    i20TestcasesL3//17

            //};

            //var g08LCITestcases = GetCarlineTestcase(reader, "SW.2041.121_G08LCI_Tracking.xlsm");
            //var g26Testcases = GetCarlineTestcase(reader, "SW.2041.200_G26_Tracking.xlsm");
            //var g28Testcases = GetCarlineTestcase(reader, "SW.2041.120_G28_Tracking.xlsm");
            //var g60Testcases = GetCarlineTestcase(reader, "SW.2149.0_G60_Tracking.xlsm");
            //var g70Testcases = GetCarlineTestcase(reader, "SW.2137.21_G70_Tracking.xlsm");
            //var i20Testcases = GetCarlineTestcase(reader, "SW.2041.121_I20_Top_Tracking.xlsm");



            //var testCasesAllCarlines = new[]
            //{
            //    g08LCITestcases,
            //    g26Testcases,
            //    g28Testcases,
            //    g60Testcases,
            //    g70Testcases,
            //    i20Testcases
            //};

            //var oldSpecTestcaseIDs = testCases.Select(t => t.ID).ToList();
            //var carlineTestCaseIDs = testCasesAllCarlines.SelectMany(t => t).Select(t => t.ID).ToList();

            //var allTestCaseIDs = oldSpecTestcaseIDs
            //    .Union(carlineTestCaseIDs)
            //    .Distinct()
            //    .Where(t => !string.IsNullOrWhiteSpace(t))
            //    .OrderBy(t => t)
            //    .ToList();



            //var tsrID_asilAB = reader.ReadFile("V15107D_FS_TechnicalSafetyRequirements_v1.08.xlsx", "CCU TSR List", 22, t => new Eng9_TSR(t))
            //    .Where(t => t.TSRID != null)
            //    .Where(t => t.ASIL != null)
            //    .Where(t => t.ASIL.Contains('A') || t.ASIL.Contains('B'))
            //    .ToList();

            //foreach (var tsr in tsrID_asilAB)
            //{
            //    Console.WriteLine(tsr.TSRID);

            //}

            //var testCases_eng9 = reader.ReadFile("BMW03_V15107D_SYI_TestReport_G70_2137.10.xlsx", "Outline of FuSa Test Cases", 7, t => new Eng9_TestCases(t))
            //    .Where(t => t.TSRID != null)
            //    .ToList();

            //foreach (var tc in testCases_eng9)
            //{
            //    Console.WriteLine(tc.TSRID);
            //    Console.WriteLine(tc.TestCaseID);

            //}


            //new HtmlReportGenerator().GenerateReport(testCases, requirements);
            //new ExcelsReportGenerator().GenerateReport(testCases, requirements);
            //new ExcelReportWithAWorkbookGenerator().GenerateReport(testCases, requirements);

            //var spec = new SpecParameters(
            //    requirements,
            //    testCases,
            //    delReqs,
            //    rejReqs,
            //    null,
            //    allCarlineTestcaseDetails,
            //    allTestCaseIDs);

            //new ExcelReportWithAWorkbookGenerator().GenerateTestSpec(spec);


            //Checker aChecker = new Checker(testCases, requirements, delReqs, rejReqs);

            //HashSet<int> usedRequirement = aChecker.FindUsedRequirments();

            //HashSet<int> unusedRequirement = aChecker.FindUnusedRequirments();
            //HashSet<int> deletedRequirementsinTC = aChecker.FindDeletedRequirements();
            //HashSet<string> delTestcases = aChecker.FindDeletedTestcases();
            //HashSet<string> rejTestcases = aChecker.FindRejectedTestcases();

            //foreach (var rejTC in rejTestcases)
            // {
            //     foreach(var tc in testCases)
            //     {
            //         if(rejTC == tc.ID)
            //         {
            //             Console.WriteLine($"{tc.ID} {tc.TotalTestResults} {tc.TestResultFuSi} {tc.TestResultFunctional}");
            //             break;
            //         }
            //     }
            // }

            //new ExcelReportWithAWorkbookGenerator().GenerateTestSpec(testCases, requirements);

            //foreach (var tc in testCases_eng9)
            //{
            //    foreach (var tsr in tsrID_asilAB)
            //    {

            //        if(tsr.TSRID == tc.TSRID)
            //        {
            //            Console.WriteLine($"[{tsr.TSRID}],[{tc.TestCaseID}]");
            //            //Console.WriteLine(tc.TestCaseID);
            //            break;

            //        }

            //    }
            //}

        }

        

        //private static List<CarlineTestcases> GetCarlineTestcase(DataFileReader reader, string file)
        //{
        //    return reader.ReadFile(file, "Summary", 4, t => new CarlineTestcases(t)).ToList();
        //}
    }
}
