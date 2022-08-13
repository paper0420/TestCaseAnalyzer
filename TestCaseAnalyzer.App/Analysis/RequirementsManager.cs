using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCaseAnalyzer.App
{
    public static class RequirementsManager
    {
        public static void CompareBaseline(SpecParameters spec)
        {
            HashSet<string> oReq = new HashSet<string>();
            HashSet<string> nReq = new HashSet<string>();

            foreach (var oldReq in spec.CurrentRequirements)
            {
                oReq.Add(oldReq.ID);
            }

            foreach (var newReq in spec.NewRequirements)
            {
                nReq.Add(newReq.ID);
            }

            foreach (var aNewReq in nReq)
            {
                if (!oReq.Contains(aNewReq))
                {
                    //Console.WriteLine($"{aNewReq}");
                    continue;
                }
            }


            foreach (var oldReq in spec.CurrentRequirements)
            {
                foreach (var newReq in spec.NewRequirements)
                {
                    if (oldReq.ID == newReq.ID)
                    {
                        if (oldReq.Objective != newReq.Objective ||
                            oldReq.panaStatus != newReq.panaStatus ||
                            oldReq.changeStatus != newReq.changeStatus ||
                            oldReq.VerificationMeasure != newReq.VerificationMeasure)
                        {
                            Console.WriteLine($"{oldReq.ID}|Changed");
                            break;

                        }
                        else
                        {
                            Console.WriteLine($"{oldReq.ID}|Same");
                            break;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
            }

        }

        public static void FindTestCasesLinked(SpecParameters spec)
        {
            foreach (var syr in spec.syrItems)
            {
                foreach (var req in syr.RequirementIDs)
                {
                    var detail = $"{req}";
                    var status = "";
                    var tcIDLinked = "";
                    var verificationMeasure = "";


                    var isDeleted = spec.delReqs.Any(t => t.ID == req);
                    if (isDeleted)
                    {
                        status = "Deleted";
                    }

                    if (!isDeleted)
                    {
                        var isRejected = spec.rejReqs.Any(t => t.ID == req);

                        if (isRejected)
                        {
                            status = "Rejected";
                        }
                    }

                    var isLinked = false;
                    foreach (var tc in spec.TestCases)
                    {
                        var isTCLinked = tc.RequirementIDs.Any(t => t == req);
                        if (isTCLinked)
                        {
                            tcIDLinked += $"{tc.ID} ";
                            isLinked = true;
                        }

                    }

                    if (!isLinked)
                    {
                        tcIDLinked = "No";
                    }

                    foreach (var KLH in spec.CurrentRequirements)
                    {
                        if (req == KLH.ID)
                        {
                            verificationMeasure = KLH.VerificationMeasure;
                            break;
                        }
                    }

                    Console.WriteLine($"{syr.ID}|{syr.Objective}|{detail}|{status}|{tcIDLinked}|{syr.Allocation}|{verificationMeasure}|{syr.SYR_ID}");

                }
            }

        }

        public static void FindENG9TestCasesLinked(SpecParameters spec)
        {
            foreach (var syr in spec.syrLists)
            {
                var testCaseID = "";
                var isLinked = false;
                foreach (var tc in spec.eng9FuncTestCases)
                {
                    if(tc.SYR_ID != null)
                    {
                        if (syr.SYR_ID.Contains(tc.SYR_ID))
                        {
                            testCaseID += $"{tc.ID} ";
                            isLinked = true;
                            continue;
                            
                        }

                    }
             

                }

                if (!isLinked)
                {
                    testCaseID = "No";
                }

                Console.WriteLine($"{syr.SYR_ID}|{testCaseID}|");

            }
        }

        public static void CheckAttributes(SpecParameters spec)
        {
            foreach(var tc in spec.TestCases)
            {
                foreach(var tcReq in tc.RequirementIDs)
                {
                    foreach(var req in spec.NewRequirements)
                    {
                        if(tcReq == req.ID)
                        {
                            if(req.VerificationMeasure != null)
                            {
                                if (!req.VerificationMeasure.Contains("SYQT (ENG.10)"))
                                {
                                    Console.WriteLine($"{tc.ID}|{req.ID}|{req.Type}|{req.VerificationMeasure}|{tc.Result}");
                                    break;
                                }

                            }
                          

                        }
                    }

                }
            }


        }
    
        public static void FindFTTtstCasesinENG10(SpecParameters spec)
        {

            foreach(var tc in spec.TestCases)
            {
                string obj = "";
                if (tc.Objective != null)
                {
                    if (tc.Objective.Contains("FTT"))
                    {
                        obj = tc.Objective.Replace("\n", " ");
                        Console.WriteLine($"{tc.ID}|{obj}");
                        continue;
                    }

                }
               
            }

            HashSet<string> klhChecked = new HashSet<string>();

            foreach (var tc in spec.TestCases)
            {
               foreach(var klh in tc.RequirementIDs)
                {
                    string obj = "";
                    if (!klhChecked.Contains(klh))
                    {
                        foreach (var req in spec.CurrentRequirements)
                        {
                            if (klh == req.ID)
                            {
                                if (req.Objective!=null)
                                {
                                    if (req.Objective.Contains("FTT"))
                                    {
                                        obj = req.Objective.Replace("\n", " ");
                                        Console.WriteLine($"{req.ID}|{obj}|{tc.ID}");
                                        klhChecked.Add(req.ID);
                                        break;
                                    }

                                }
                               
                            }
                        }

                    }
                    
                }

            }



        }
    
    }


}

