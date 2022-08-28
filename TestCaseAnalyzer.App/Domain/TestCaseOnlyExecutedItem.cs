﻿using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;
using TestCaseAnalyzer.App.FileReader;

namespace TestCaseAnalyzer.App
{
    public class TestCaseOnlyExecutedItem
    {
        private static IEnumerable<string> GetAvailableCarlines(
            IExcelDataReader reader,
            Header header,
            List<string> carLineNames)
        {
            foreach (var carLineName in carLineNames)
            {
                var carLineColumnName = $"#{carLineName}#";
                var columnIndex = header.GetColumnIndex(carLineColumnName);
            
                var contains = reader.GetValue(columnIndex)?.ToString()?.Contains("X") == true;

                if (contains)
                {
                    yield return carLineName;
                }
            }
        }

        internal static TestCaseOnlyExecutedItem? CreateOrNull(IExcelDataReader reader, Header header)
        {
            var id = reader.GetValue(header.GetColumnIndex("Test Case ID"))?.ToString();

            if (string.IsNullOrWhiteSpace(id))
            {
                return null;
            }

            var result = new TestCaseOnlyExecutedItem();

            result.ID = id;
            result.Objective = reader.GetString(header.GetColumnIndex("Test objective"));
            var idsAsString = reader.GetValue(header.GetColumnIndex("Current KLH"))?
                .ToString()?
                .Split('\n') ?? new string[0];

            //var requirementIds = new List<int>();
            var requirementIds = new List<string>();
            var epicIds = new List<string>();


            foreach (var idAsString in idsAsString)
            {
                //var isInt = int.TryParse(idAsString, out int idAsInt);

                //if (isInt)
                //{
                //    requirementIds.Add(idAsInt);
                //}
                //else
                //{
                //    epicIds.Add(idAsString);
                //}

                requirementIds.Add(idAsString);


            }

            result.RequirementIDs = requirementIds.ToArray();
            result.EpicIDs = epicIds.ToArray();

            if (CarLineNames.carLineNames == null)
            {
                CarLineNames.carLineNames = header.Columns
                .Where(t => t.Name != null)
                .Where(t => t.Name!.StartsWith("#"))
                .Where(t => t.Name!.EndsWith("#"))
                .Select(t => t.Name!.Replace("#", "").Replace("#", ""))
                .ToList();
            }

            result.Carlines = GetAvailableCarlines(
                    reader,
                    header,
                    CarLineNames.carLineNames)
                .ToList();

            result.Type = reader.GetValue(header.GetColumnIndex("Type"))?.ToString();
            result.Result = reader.GetValue(header.GetColumnIndex("Result"))?.ToString()?.Replace("\n", " ");

            result.ItemClass1 = reader.GetString(header.GetColumnIndex("Class1"));
            result.ItemClass2 = reader.GetString(header.GetColumnIndex("Class2"));
            result.ItemClass3 = reader.GetString(header.GetColumnIndex("Class3"));

            result.Comment = reader.GetValue(header.GetColumnIndex("Comment"))?.ToString();

            result.VerificationMethod = reader.GetValue(header.GetColumnIndex("Verification Method"))?.ToString();
            result.TestCatHV = reader.GetValue(header.GetColumnIndex("TestCat-HV"))?.ToString();
            result.TestCatBasic = reader.GetValue(header.GetColumnIndex("TestCat-HV"))?.ToString();
            result.TestCatFusa = reader.GetValue(header.GetColumnIndex("TestCat-Fusa"))?.ToString();
            result.TestCatFunc = reader.GetValue(header.GetColumnIndex("TestCat-Func"))?.ToString();
            result.TestCatFull = reader.GetValue(header.GetColumnIndex("TestCat-Full"))?.ToString();

            return result;
        }

        public string ID { get; private set; }
        public string[] RequirementIDs { get; private set; }
        public string[] EpicIDs { get; private set; }
        public List<string> Carlines { get; private set; }
        public string Type { get; private set; }
        public string Result { get; private set; }
        public string Objective { get; private set; }
        public string ItemClass1 { get; private set; }
        public string ItemClass2 { get; private set; }
        public string ItemClass3 { get; private set; }
        public string Comment { get; private set; }
        public string VerificationMethod { get; private set; }
        public string TestCatHV { get; private set; }
        public string TestCatBasic { get; private set; }
        public string TestCatFusa { get; private set; }
        public string TestCatFunc { get; private set; }
        public string TestCatFull { get; private set; }
    }
}
