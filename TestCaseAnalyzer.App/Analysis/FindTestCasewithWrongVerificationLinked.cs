using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace TestCaseAnalyzer.App.Analysis
{
    public static class FindTestCasewithWrongVerificationLinked
    {
        public static void FindTCwWrongVeriMea(SpecParameters spec)
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("verificationMeasure");
            int currentRow = 1;
            foreach(var specTC in spec.TestCases)
            {
                var klhs = "";
                if (specTC.Result != null)
                {
                    
                    foreach (var req in specTC.RequirementIDs)
                    {
                        var klhID = spec.CurrentRequirementsByID.ContainsKey(req)
                          ? spec.CurrentRequirementsByID[req]
                          : null;

                        if (klhID == null)
                        {
                            continue;
                        }

                        if (!klhID.VerificationMeasure.Contains("ENG.10"))
                        {
                            klhs += $"{klhID.ID}\n";
                            

                        }
                    }

                    if(klhs == "")
                    {
                        continue;
                    }
                    else
                    {
                        ws.Cell($"A{currentRow}").Value = specTC.ID;
                        ws.Cell($"B{currentRow}").Value = klhs;
                        currentRow++;
                    }


                }
            }
            wb.SaveAs("wrongVerificationTCs.xlsx");
        }
    }
}
