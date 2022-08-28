using System;
using System.Collections.Generic;
using System.Linq;
using TestCaseAnalyzer.App.Domain;

namespace TestCaseAnalyzer.App.FileReader;

public class Header
{
    private Dictionary<string,Column> columnsById;
    public List<Column> Columns { get; } = new();
    
    public int GetColumnIndex(string name)
    {
        if (this.columnsById == null)
        {
            this.columnsById = this.Columns
                .Where(t => !string.IsNullOrWhiteSpace(t.Name))
                .ToDictionary(t => t.Name);
        }

        if (this.columnsById.ContainsKey(name))
        {
            return this.columnsById[name].Index;    
        }

        throw new Exception($"Cannot find column `{name}`.");
    }

    public List<string> GetCarLineNames(List<Column> columns)
    {
        List<string> carLines = new List<string>();
        foreach (Column column in columns)
        {

            if (column.Name != null && column.Name.Contains("#"))
            {
                carLines.Add(column.Name);
                continue;
            }

        }
        return CarLineNames.carLineNames = carLines;

    }


}