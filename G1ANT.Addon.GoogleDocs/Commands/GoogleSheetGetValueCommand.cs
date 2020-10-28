/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.GoogleDocs
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/

using G1ANT.Language;
using Google.Apis.Sheets.v4.Data;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace G1ANT.Addon.GoogleDocs
{
    [Command(Name = "googlesheet.getvalue", Tooltip = "This command gets a value from a specified cell or range in an opened Google Sheets instance")]
    public class GoogleSheetGetValueCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Cell or range to get a value from, for example `A6`, `A3&B7` or `A3:A6&B7`")]
            public TextStructure Range { get; set; }

            [Argument(Tooltip = "Sheet name containing the specified range; can be empty or omitted")]
            public TextStructure SheetName { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public GoogleSheetGetValueCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            var sheetName = arguments.SheetName.Value == "" ? sheetsManager.Sheets[0].Properties.Title : arguments.SheetName.Value;
            var val = sheetsManager.GetValue(arguments.Range.Value, sheetName);


            var groups = new List<DataTable>(); // groups of columns of cells

            if (val.ValueRanges[0].Values != null)
            {
                groups = val.ValueRanges.Select(group => WrapGroupInDataTable(group)).ToList();
            }

            var resultStructure = new ListStructure(groups.Select(g => new DataTableStructure(g)));

            Scripter.Variables.SetVariableValue(arguments.Result.Value, resultStructure);
        }

        private DataTable WrapGroupInDataTable(ValueRange group)
        {
            var result = new DataTable(group.Range);
            var maxColumnCount = group.Values.Max(r => r.Count);

            result.Columns.AddRange(
                Enumerable.Range(0, maxColumnCount).Select(c => new DataColumn(c.ToString())).ToArray()
            );

            foreach (var row in group.Values)
            {
                var dataTableRow = result.NewRow();
                dataTableRow.ItemArray = row.ToArray();
                result.Rows.Add(dataTableRow);
            }

            return result;
        }

    }
}
