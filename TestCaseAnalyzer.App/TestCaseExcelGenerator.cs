using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;

namespace TestCaseAnalyzer.App
{
    public class TestCaseExcelGenerator
    {
        public static void CreateTestSpecExcel(SpecParameters spec)
        {
    

            Console.WriteLine("write Data");

            List<string> testCaseID = new List<string>();
            HashSet<string> newReqs = new HashSet<string>();
            HashSet<string> delRejReqs = new HashSet<string>();
            List<int> carLineNumbers = new List<int>();
            HashSet<string> checkDuplicated = new HashSet<string>();


            int row = 1;

            spec.xlsSheet["A1"].Value = "Test Case ID";
            spec.xlsSheet["B1"].Value = "Class1";
            spec.xlsSheet["C1"].Value = "Class2";
            spec.xlsSheet["D1"].Value = "Class3";
            spec.xlsSheet["E1"].Value = "Test objective";
            spec.xlsSheet["F1"].Value = "Old KLH";
            spec.xlsSheet["G1"].Value = "Current KLH";
            spec.xlsSheet["H1"].Value = "Deleted/Rejected KLH";
            spec.xlsSheet["I1"].Value = "Type";
            spec.xlsSheet["J1"].Value = "G08LCI";
            spec.xlsSheet["K1"].Value = "G26";
            spec.xlsSheet["L1"].Value = "G28";
            spec.xlsSheet["M1"].Value = "G60";
            spec.xlsSheet["N1"].Value = "G70";
            spec.xlsSheet["O1"].Value = "I20";
            spec.xlsSheet["P1"].Value = "Test Time";
            spec.xlsSheet["Q1"].Value = "Result";
            spec.xlsSheet["R1"].Value = "Ticket Number";
            spec.xlsSheet["S1"].Value = "Comment";
            spec.xlsSheet["T1"].Value = "Param Sheet";


            foreach (var testCase in spec.allTestCaseIDs)
            {
                int carline = 0;
                carLineNumbers.Clear();
                string concatEpic ="";
                //string testTime = "";
                //string result = "";
                //string comment = "";
                //string ticketNumber = "";

                //if (testCase.ID != null)
                {
                    //if (!testCaseID.Contains(testCase.ID))
                    {
                        for (int column = 0; column <= 25; column++)
                        {
                            switch (column)
                            {
                                case 0:
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.ID}");
                                            break;
                                        }
                                        else
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{testCase}");
                                        }
                                    }

                                    break;

                                case 1:
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.ItemClass1}");
                                            break;
                                        }
                                    }

                                    break;
                                case 2:
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.ItemClass2}");
                                            break;
                                        }
                                    }

                                    break;
                                case 3:
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.ItemClass3}");
                                            break;
                                        }
                                    }

                                    break;
                                case 4:
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.Objective}");
                                            break;
                                        }
                                    }

                                    break;

                                case 5: //Old KLH
                                    
                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            string concatReqs = String.Join("\n", tc.RequirementIDs.ToArray());
                                            if(tc.EpicIDs.Length != 0)
                                            {
                                                concatEpic = String.Join("\n", tc.EpicIDs.ToArray());
                                                spec.xlsSheet.SetCellValue(row, column, $"{concatReqs}\n{concatEpic}");
                                            }
                                            else
                                            {
                                                spec.xlsSheet.SetCellValue(row, column, $"{concatReqs}");
                                            }
                                            
                              
                                        }
                                    }

                                    break;

                                case 6:

                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {

                                            foreach (var tcReq in tc.RequirementIDs)
                                            {
                                                foreach (var dlreq in spec.delReqs)
                                                {
                                                    if (dlreq.ID == tcReq)
                                                    {
                                                        delRejReqs.Add(tcReq);
                                                        break;
                                                    }

                                                }
                                            }


                                            foreach (var tcReq in tc.RequirementIDs)
                                            {
                                                foreach (var rjreq in spec.rejReqs)
                                                {
                                                    if (rjreq.ID == tcReq)
                                                    {
                                                        delRejReqs.Add(tcReq);
                                                        break;

                                                    }
                                                }
                                            }


                                            foreach (var tcReq in tc.RequirementIDs)
                                            {
                                                if (delRejReqs.Any(d => d == tcReq))
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    newReqs.Add(tcReq);
                                                }
                                            }

                                            //string concatEpic = String.Join("\n", tc.EpicIDs.ToArray());
                                            string newReq = String.Join("\n", newReqs.ToArray());
                                            if (tc.EpicIDs.Length != 0)
                                            {
                                                concatEpic = String.Join("\n", tc.EpicIDs.ToArray());
                                                spec.xlsSheet.SetCellValue(row, column, $"{newReq}\n{concatEpic}");
                                            }
                                            else
                                            {
                                                spec.xlsSheet.SetCellValue(row, column, $"{newReq}");
                                            }
                                           

                                        }
                                    }


                                    newReqs.Clear();


                                    break;

                                case 7: //deleted rejected KLH

                                    string delRejReq = String.Join("\n", delRejReqs.ToArray());
                                    {
                                        spec.xlsSheet.SetCellValue(row, column, $"{delRejReq}");
                                    }

                                    delRejReqs.Clear();
                                    break;

                                case 8: // Type L1 L2 L3

                                    for (var ind = 0;ind<=5;ind++)
                                    {
                                        
                                        foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                        {
                                            if (testCase==carlineDetail.ID)
                                            {
                                                spec.xlsSheet.SetCellValue(row, column, $"L1");
                                                carLineNumbers.Add(ind);
                                                carline = 1;
                                              
                                            }
                                        }
                                    }

                                    if (carline == 1)
                                    {
                                        break;
                                    }

                                    if (carline == 0)
                                    {
                                        carLineNumbers.Clear();
                                        for (var ind = 6; ind <= 11; ind++)
                                        {
                                            
                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    spec.xlsSheet.SetCellValue(row, column, $"L2");
                                                    carLineNumbers.Add(ind);
                                                    carline = 2;
                                                    
                                                }
                                            }
                                        }
                                    }

                                    if (carline == 2)
                                    {
                                        break;
                                    }

                                    if (carline == 0)
                                    {
                                        carLineNumbers.Clear();
                                        for (var ind = 12; ind <= 17; ind++)
                                        {
                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    spec.xlsSheet.SetCellValue(row, column, $"L3");
                                                    carLineNumbers.Add(ind);
                                                    carline = 3;
                                                   
                                                }
                                            }
                                        }
                                    }
                                    


                                    break;

                                case 9: //G08LCI 

                                    //string carLineIndex = String.Join(" ", carLineNumbers.ToArray());
                                    //{
                                    //    spec.xlsSheet.SetCellValue(row, column, $"{carLineIndex}");
                                    //}

                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 0 || carLinenumber == 6 || carLinenumber == 12)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }


                                    break;

                                case 10: //G26
                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 1 || carLinenumber == 7 || carLinenumber == 13)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }
                                    break;

                                case 11://G28
                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 2 || carLinenumber == 8 || carLinenumber == 14)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }
                                    break;


                                case 12://G60
                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 3 || carLinenumber == 9 || carLinenumber == 15)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }
                                    break;


                                case 13://G70
                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 4 || carLinenumber == 10 || carLinenumber == 16)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }
                                    break;

                                case 14://I20
                                    foreach (var carLinenumber in carLineNumbers)
                                    {
                                        if (carLinenumber == 5 || carLinenumber == 11 || carLinenumber == 17)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"X");
                                        }
                                    }
                                    break;


                                case 15: //Test time
                                    if (carline == 1)
                                    {
                                        
                                        for (var ind = 0; ind <= 5; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    spec.xlsSheet.SetCellValue(row, column, $"{carlineDetail.TestTime}");
                                                    break;

                                                }
                                            }
                                        }

                                       
                                        break;

                                    }

                                    if (carline == 2)
                                    {
                                        
                                        for (var ind = 6; ind <= 11; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    spec.xlsSheet.SetCellValue(row, column, $"{carlineDetail.TestTime}");
                                                    break;

                                                }
                                            }
                                        }

                         
                                        break;

                                    }

                                    if (carline == 3)
                                    {
                                        
                                        for (var ind = 12; ind <= 17; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    spec.xlsSheet.SetCellValue(row, column, $"{carlineDetail.TestTime}");
                                                    break;

                                                }
                                            }
                                        }


                                    }

                                    break;

                                case 16: //History
                                    if (carline == 1)
                                    {
                                        string detail = "";
                                        int count = 1;
                                        string name = "";
                                        for (var ind = 0; ind <= 5; ind++)
                                        {
                                            
                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    name = GetCarlineName(count);
                                                    detail += $"[{name}] {carlineDetail.History}\n";

                                                }
                                            }
                                            count++;
                                        }

                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 2)
                                    {
                                        string detail = "";
                                        int count = 1;
                                        string name = "";
                                        for (var ind = 6; ind <= 11; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    name = GetCarlineName(count);
                                                    detail += $"[{name}] {carlineDetail.History}\n";

                                                }
                                            }
                                            count++;
                                        }

                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 3)
                                    {
                                        string detail = "";
                                        int count = 1;
                                        string name = "";
                                        for (var ind = 12; ind <= 17; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    name = GetCarlineName(count);
                                                    detail += $"[{name}] {carlineDetail.History}\n";

                                                }
                                            }
                                            count++;

                                        }

                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    break;

                                case 17: //TicketNumber
                                    if (carline == 1)
                                    {
                                        string detail = "";
                                        int count = 1;
                              
                                        for (var ind = 0; ind <= 5; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    if (!checkDuplicated.Contains(carlineDetail.TicketNumber))
                                                    {
                                                        detail += $"{carlineDetail.TicketNumber}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.TicketNumber);

                                                }
                                            }
                                            count++;
                                        }
                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 2)
                                    {
                                        string detail = "";
                                        int count = 1;
                                 
                                        for (var ind = 6; ind <= 11; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {

                                                    if (!checkDuplicated.Contains(carlineDetail.TicketNumber))
                                                    {
                                                        detail += $"{carlineDetail.TicketNumber}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.TicketNumber);


                                                }
                                            }
                                            count++;
                                        }
                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 3)
                                    {
                                        string detail = "";
                                        int count = 1;
                                      
                                        for (var ind = 12; ind <= 17; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {

                                                    if (!checkDuplicated.Contains(carlineDetail.TicketNumber))
                                                    {
                                                        detail += $"{carlineDetail.TicketNumber}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.TicketNumber);


                                                }
                                            }
                                            count++;
                                        }
                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    break;

                                case 18:// Comment

                                    
                                    if (carline == 1)
                                    {
                                        
                                        string detail = "";
                                        int count = 1;
                           
                                        for (var ind = 0; ind <= 5; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    if (!checkDuplicated.Contains(carlineDetail.Comment))
                                                    {
                                                        detail += $"{carlineDetail.Comment}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.Comment); 

                                                 }
  
                                                
                                            }
                                            count++;
                                        }

                                       

                                        foreach (var tc in spec.testCases)
                                        {
                                            if (testCase == tc.ID)
                                            {
                                                detail += $"{tc.Comment}";

                                            }
                                        }

                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 2)
                                    {
                                     
                                        string detail = "";
                                        int count = 1;

                                        for (var ind = 6; ind <= 11; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    if (!checkDuplicated.Contains(carlineDetail.Comment))
                                                    {
                                                        detail += $"{carlineDetail.Comment}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.Comment);

                                                }

                                            }
                                            count++;
                                        }



                                        foreach (var tc in spec.testCases)
                                        {
                                            if (testCase == tc.ID)
                                            {
                                                detail += $"{tc.Comment}";

                                            }
                                        }

                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    if (carline == 3)
                                    {
                               
                                        string detail = "";
                                        int count = 1;
                           
                                        for (var ind = 12; ind <= 17; ind++)
                                        {

                                            foreach (var carlineDetail in spec.allCarlineTestcaseDetails[ind])
                                            {
                                                if (testCase == carlineDetail.ID)
                                                {
                                                    if (!checkDuplicated.Contains(carlineDetail.Comment))
                                                    {
                                                        detail += $"{carlineDetail.Comment}\n";
                                                    }
                                                    checkDuplicated.Add(carlineDetail.Comment);

                                                }

                                            }
                                            count++;
                                        }


                                        foreach (var tc in spec.testCases)
                                        {
                                            if (testCase == tc.ID)
                                            {
                                                detail += $"{tc.Comment}";

                                            }
                                        }

                                        checkDuplicated.Clear();
                                        spec.xlsSheet.SetCellValue(row, column, $"{detail}");
                                        break;

                                    }

                                    foreach (var tc in spec.testCases)
                                    {
                                        if (testCase == tc.ID)
                                        {
                                            spec.xlsSheet.SetCellValue(row, column, $"{tc.Comment}");
                                         
                                        }
                                    }

                                    break;



                            }

                        }

                        row++;
                    }
                }

                //testCaseID.Add(testCase.ID);
            }

        }


        public static void CreateTestCaseExcel(
            List<Requirement> requirements,
            TestCase testCase,
            WorkSheet xlsSheet)
        {
            GenerateRequirementsExcel(
                requirements,
                testCase,
                xlsSheet);

            xlsSheet["A1"].Value = "Test Case ID";
            xlsSheet["C1"].Value = $"{testCase.ID}";
            xlsSheet["D1"].Value = $"[{testCase.TotalTestResults}],[{testCase.TestResultFuSi}],[{testCase.TestResultFunctional}]";
            xlsSheet["A2"].Value = "Test Objective";
            xlsSheet["C2"].Value = $"{testCase.Objective}";
            xlsSheet["D2"].Value = $"[{testCase.ItemClass1}],[{testCase.ItemClass2}],[{testCase.ItemClass3}]";
            xlsSheet["A3"].Value = "Requirements";
        }

        private static void GenerateRequirementsExcel(List<Requirement> requirements, TestCase testCase, WorkSheet xlsSheet)
        {
            var testCaseRequirements = requirements
                .Where(t => testCase.RequirementIDs.Any(c => c == t.ID))
                .ToList();

            int row = 5;
            foreach (var requirement in testCaseRequirements)
            {


                //    xlsSheet[$"A{Row}"].Value = $"{testCase.ID}";
                //    xlsSheet[$"B{Row}"].Value = $"{testCase.Objective}";
                xlsSheet[$"C{row}"].Value = $"{requirement.ID},[{requirement.changeStatus}],[{requirement.panaStatus}]";
                xlsSheet[$"D{row}"].Value = $"{requirement.Objective}";
                //xlsSheet[$"E{Row}"].Value = $"{string.Join(", ", requirement.EpicIDs)}";
                row++;
            }


        }

        private static string GetCarlineName (int count)
        {
            string carlineName ="";
            switch (count)
            {
                case 1:
                    carlineName = "G80LCI";
                    break;

                case 2:
                    carlineName = "G26";
                    break;

                case 3:
                    carlineName = "G28";
                    break;
                case 4:
                    carlineName = "G60";
                    break;

                case 5:
                    carlineName = "G70";
                    break;

                case 6:
                    carlineName = "I20";
                    break;

            }
            return carlineName;
        }




    }
}
