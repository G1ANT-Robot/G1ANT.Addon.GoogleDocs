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

    [Command(Name = "googlesheet.getrow", Tooltip = "This command returns all data from a row")]
    public class GoogleSheetGetRowCommand : Command
    {

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Row index (e.g. `1`) which will be returned by this command")]
            public IntegerStructure Row { get; set; }

            [Argument(Tooltip = "Sheet name; can be empty or omitted")]
            public TextStructure SheetName { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            
        }
        
        public GoogleSheetGetRowCommand(AbstractScripter scripter) : base(scripter)
        { }

        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            var sheetName = arguments.SheetName.IsNullOrEmpty() ? sheetsManager.Sheets[0].Properties.Title : arguments.SheetName.Value;
            var row = sheetsManager.GetRow(sheetName, arguments.Row.Value);
            
            var result = new ListStructure(row.Select(c => new TextStructure(c)));
            Scripter.Variables.SetVariableValue(arguments.Result.Value, result);
        }

    }
}
