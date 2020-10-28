/**
*    Copyright(C) G1ANT Ltd, All rights reserved
*    Solution G1ANT.Addon, Project G1ANT.Addon.GoogleDocs
*    www.g1ant.com
*
*    Licensed under the G1ANT license.
*    See License.txt file in the project root for full license information.
*
*/
using G1ANT.Addon.GoogleDocs.Helpers;
using G1ANT.Language;

namespace G1ANT.Addon.GoogleDocs
{

    [Command(Name = "googlesheet.getrowcount", Tooltip = "This command returns number of rows in a sheet")]
    public class GoogleSheetGetRowCountCommand : Command
    {

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Column name (e.g. `A`) for which the rows will be counted. Column name is necessary, because number of rows may vary between the columns")]
            public TextStructure Column { get; set; }

            [Argument(Tooltip = "Sheet name; can be empty or omitted")]
            public TextStructure SheetName { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            
        }
        
        public GoogleSheetGetRowCountCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            var sheetName = arguments.SheetName.IsNullOrEmpty() ? sheetsManager.Sheets[0].Properties.Title : arguments.SheetName.Value;
            var result = sheetsManager.GetRowCount(sheetName, arguments.Column.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new IntegerStructure(result));
        }

    }
}
