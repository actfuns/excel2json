using CommandLine;

namespace ActFuns.Tools.Excel2Json
{
    partial class Program
    {
        /// <summary>
        /// 命令行参数定义
        /// </summary>
        internal sealed class Options
        {
            /// <summary>
            /// ExcelPath
            /// </summary>
            [Option('e', "excel", Required = true, HelpText = "input excel file path.")]
            public string ExcelPath { get; set; }

            /// <summary>
            /// JsonPath
            /// </summary>
            [Option('j', "json", Required = false, HelpText = "export json file path.")]
            public string JsonPath { get; set; }

            /// <summary>
            /// HeaderRows
            /// </summary>
            [Option('h', "header", Required = false, Default = 1, HelpText = "number lines in sheet as header.")]
            public int HeaderRows { get; set; }

            /// <summary>
            /// Encoding
            /// </summary>
            [Option('c', "encoding", Required = false, Default = "utf8-nobom", HelpText = "export file encoding.")]
            public string Encoding { get; set; }

            /// <summary>
            /// Sheet
            /// </summary>
            [Option('s', "sheet", Required = false, HelpText = "input excel sheet name.")]
            public string Sheet { get; set; }

            /// <summary>
            /// ExportArray
            /// </summary>
            [Option('a', "array", Required = false, Default = false, HelpText = "export as array, otherwise as dict object.")]
            public bool ExportArray { get; set; }

            /// <summary>
            /// DateFormat
            /// </summary>
            [Option('d', "date", Required = false, Default = "yyyy/MM/dd", HelpText = "Date Format String, example: dd / MM / yyy hh: mm:ss.")]
            public string DateFormat { get; set; }
        }
    }
}
