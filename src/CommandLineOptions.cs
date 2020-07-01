using CommandLine;

namespace VandanaMathure.DemoExcelUpload
{
    /// <summary>
    /// Represents commandline options 
    /// </summary>
    internal sealed class CommandLineOptions
    {
        /// <summary>
        /// Gets/Sets full name of excel file to be processed
        /// </summary>
        [Option('f', "FileName", HelpText = "Full name of excel file to be converted", Required = true)]
        public string FileName { get; set; }

        /// <summary>
        /// Gets/Sets name of table
        /// </summary>
        [Option('w', "WorksheetName", HelpText = "Name of excel worksheet", Required = true)]
        public string WorksheetName { get; set; }

        /// <summary>
        /// Gets/Sets Start Row Index
        /// </summary>
        [Option('s', "StartRowIndex", HelpText = "One based index of start row. This is row number in excel sheet", Required = true)]
        public int StartRowIndex { get; set; }

        /// <summary>
        /// Gets/Sets End Row Index
        /// </summary>
        [Option('e', "EndRowIndex", HelpText = "One based index of end row. This is row number in excel sheet", Required = true)]
        public int EndRowIndex { get; set; }

        /// <summary>
        /// Gets/Sets Start Column Index
        /// </summary>
        [Option('l', "StartColumnIndex", HelpText = "One based index of start column. Column A in excel becomes 1, Column B becomes 2 and so on", Required = true)]
        public int StartColumnIndex { get; set; }

        /// <summary>
        /// Gets/Sets End Column Index
        /// </summary>
        [Option('h', "EndColumnIndex", HelpText = "One based index of end column. Column A in excel becomes 1, Column B becomes 2 and so on", Required = true)]
        public int EndColumnIndex { get; set; }
    }
}
