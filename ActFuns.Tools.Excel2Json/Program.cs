using System;
using System.Data;
using System.IO;
using System.Text;
using CommandLine;

namespace ActFuns.Tools.Excel2Json
{
    partial class Program
    {
        /// <summary>
        /// Main
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(opts => run(opts));
        }

        /// <summary>
        /// run
        /// </summary>
        /// <param name="options"></param>
        private static void run(Options options)
        {
            try
            {
                Encoding cd = new UTF8Encoding(false);
                if (options.Encoding != "utf8-nobom")
                {
                    foreach (EncodingInfo ei in Encoding.GetEncodings())
                    {
                        Encoding e = ei.GetEncoding();
                        if (e.HeaderName == options.Encoding)
                        {
                            cd = e;
                            break;
                        }
                    }
                }
                options.Encoding = cd.EncodingName;
                if (!File.Exists(options.ExcelPath))
                {
                    Console.WriteLine($"excel文件路径有误!({options.ExcelPath})");
                    Environment.Exit(-1);
                    return;
                }
                if (!string.IsNullOrEmpty(options.JsonPath))
                {
                    string dir = Path.GetDirectoryName(options.JsonPath);
                    dir = dir.Contains(":") ? dir : Path.Combine(Environment.CurrentDirectory, dir);
                    if (!Directory.Exists(dir))
                    {
                        Console.WriteLine($"json文件路径有误!({options.JsonPath})");
                        Environment.Exit(-1);
                        return;
                    }
                }
                else
                {
                    options.JsonPath = Path.ChangeExtension(options.ExcelPath, ".json");
                }
                var excelLoader = new ExcelLoader(options.ExcelPath);
                DataTable table = excelLoader.Sheets[0];
                if (!string.IsNullOrEmpty(options.Sheet))
                {
                    foreach (DataTable tbl in excelLoader.Sheets)
                    {
                        if (tbl.TableName == options.Sheet)
                        {
                            table = tbl;
                            break;
                        }
                    }
                }
                var removeNums = options.HeaderRows - 1;
                while (removeNums-- > 0 && table.Rows.Count > 0)
                {
                    table.Rows.RemoveAt(0);
                }
                var jsonExporter = new JsonExporter(table);
                jsonExporter.SaveToFile(options.JsonPath, cd, options.ExportArray, options.DateFormat);
            }
            catch (Exception exp)
            {
                Console.WriteLine("Error: " + exp.Message);
            }
        }
    }
}
