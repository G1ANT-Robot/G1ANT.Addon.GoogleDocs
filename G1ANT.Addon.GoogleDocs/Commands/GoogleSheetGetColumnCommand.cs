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
using System.Linq;

namespace G1ANT.Addon.GoogleDocs
{

    [Command(Name = "googlesheet.getcolumn", Tooltip = "This command returns all data from a column")]
    public class GoogleSheetGetColumnCommand : Command
    {
        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Column name (e.g. `A`) which will be returned by this command")]
            public TextStructure Column { get; set; }

            [Argument(Tooltip = "Sheet name; can be empty or omitted")]
            public TextStructure SheetName { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            
        }
        
        public GoogleSheetGetColumnCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            var sheetName = arguments.SheetName.IsNullOrEmpty() ? sheetsManager.Sheets[0].Properties.Title : arguments.SheetName.Value;
            var column = sheetsManager.GetColumn(sheetName, arguments.Column.Value);
            
            var result = new ListStructure(column.Select(c => new TextStructure(c)));
            Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
        }
    }
}
