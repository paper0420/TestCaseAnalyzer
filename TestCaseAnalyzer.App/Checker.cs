using System;
using System.Collections.Generic;
using System.Linq;

namespace TestCaseAnalyzer.App
{
    public class Checker
    {
        public Checker(List<TestCase> testCases,
                       List<Requirement> requirements,
                       List<DeletedReq> delReqs,
                       List<RejectedReq> rejReqs)
        {
            TestCases = testCases;
            Requirements = requirements;
            DeletedReqs = delReqs;
            RejectedReqs = rejReqs;
        }

        public List<TestCase> TestCases { get; }
        public List<Requirement> Requirements { get; }

        public List<DeletedReq> DeletedReqs { get; }

        public List<RejectedReq> RejectedReqs { get; }

        internal HashSet<string> FindDeletedRequirements()
        {
            var deletedRequiremens = new HashSet<string>();
            foreach (var requirement in DeletedReqs)
            {

                foreach (var testCase in TestCases)
                {
                    if (testCase.RequirementIDs.Any(t => t == requirement.ID))
                    {
                        //if (!usedRequirement.Contains(requirement.ID))
                        //{
                        //    usedRequirement.Add(requirement.ID);
                        //} 

                        deletedRequiremens.Add(testCase.RequirementIDs[0]);
                    }
                }
            }

            return deletedRequiremens;

        }

        internal HashSet<string> FindDeletedTestcases()
        {
            HashSet<string> faultTestcase = new HashSet<string>();

            foreach (var testcase in TestCases)
            {
                foreach (var tcReq in testcase.RequirementIDs)
                {
                    foreach (var req in DeletedReqs)
                    {
                        if (req.ID == tcReq)
                        {
                            faultTestcase.Add(testcase.ID);

                        }
                    }
                }
            }

            return faultTestcase;

        }

        internal HashSet<string> FindRejectedTestcases()
        {
            HashSet<string> faultTestcase = new HashSet<string>();

            foreach (var testcase in TestCases)
            {
                foreach (var tcReq in testcase.RequirementIDs)
                {
                    foreach (var req in RejectedReqs)
                    {
                        if (req.ID == tcReq)
                        {
                            faultTestcase.Add(testcase.ID);

                        }
                    }
                }
            }

            return faultTestcase;

        }




        internal HashSet<string> FindUnusedRequirments()
        {
            var usedRequirement = this.FindUsedRequirments();
            var unusedRequirement = new HashSet<string>();

            foreach (var requirement in Requirements)
            {

                if (!usedRequirement.Contains(requirement.ID))
                {
                    unusedRequirement.Add(requirement.ID);
                }
            }
            return unusedRequirement;

        }

        internal HashSet<string> FindUsedRequirments()
        {

            var usedRequirement = new HashSet<string>();
            foreach (var requirement in Requirements)
            {

                foreach (var testCase in TestCases)
                {
                    if (testCase.RequirementIDs.Any(t => t == requirement.ID))
                    {
                        //if (!usedRequirement.Contains(requirement.ID))
                        //{
                        //    usedRequirement.Add(requirement.ID);
                        //} 

                        usedRequirement.Add(requirement.ID);
                    }
                }
            }

            return usedRequirement;
        }
    }
}
