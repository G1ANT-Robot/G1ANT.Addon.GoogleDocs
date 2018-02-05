﻿using G1ANT.Language;

namespace G1ANT.Addon.GoogleDocs
{
    [Command(Name = "googlesheet.settitle", Tooltip = "This command allows to set title of opened Google Sheets instance.")]
    public class GoogleSheetSetTitleCommand : Command
    {
        public class Arguments : CommandArguments
        {

            [Argument(Required = true, Tooltip = "New spreadsheet title")]
            public TextStructure Title { get; set; } 

            
        }
        public GoogleSheetSetTitleCommand(AbstractScripter scripter) : base(scripter)
        { }
        public void Execute(Arguments arguments)
        {
            var sheetsManager = SheetsManager.CurrentSheet;
            sheetsManager.SetNewTitle(arguments.Title.Value);
            
        }
    }
}