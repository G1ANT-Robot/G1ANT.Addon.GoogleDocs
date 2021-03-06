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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace G1ANT.Addon.GoogleDocs
{

    [Command(Name = "googlesheet.find", Tooltip = "This command finds a cell with a specified value")]
    public class GoogleSheetFindCommand : Command
    {

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Tooltip = "Value to be searched for")]
            public TextStructure Value { get; set; }

            [Argument(Tooltip = "Sheet name where the search is to be performed; can be empty or omitted")]
            public TextStructure SheetName { get; set; } = new TextStructure(string.Empty);

            [Argument(Tooltip = "Name of a variable where the command's result will be stored")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");

            
        }
        public GoogleSheetFindCommand(AbstractScripter scripter) : base(scripter)
        { }
        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            var sheetName = arguments.SheetName.IsNullOrEmpty() ? sheetsManager.Sheets[0].Properties.Title : arguments.SheetName.Value;
           

            var result = sheetsManager.FindFirst(arguments.Value.Value, sheetName);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.TextStructure(result));
        }

    }
}
