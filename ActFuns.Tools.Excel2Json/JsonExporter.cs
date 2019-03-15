using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ActFuns.Tools.Excel2Json
{
    /// <summary>
    /// JsonExporter
    /// </summary>
    public class JsonExporter
    {
        /// <summary>
        /// regArray 
        /// </summary>
        private Regex _regArray = new Regex(@"array\<(?<type>\w+)\>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// data
        /// </summary>
        private List<Dictionary<string, object>> _data;

        /// <summary>
        /// JsonExporter
        /// </summary>
        /// <param name="table"></param>
        public JsonExporter(DataTable table)
        {
            _data = new List<Dictionary<string, object>>();
            if (table.Rows.Count < 3)
                return;

            for (var i = 2; i < table.Rows.Count; i++)
            {
                Dictionary<string, object> item = new Dictionary<string, object>();
                for (var j = 0; j < table.Columns.Count; j++)
                {
                    string key = table.Rows[0][j].ToString();
                    string type = table.Rows[1][j].ToString();
                    object data = table.Rows[i][j];
                    object value;
                    if (_regArray.IsMatch(type))
                    {
                        string sdata = data == null ? string.Empty : data.ToString();
                        string atype = _regArray.Match(type).Groups["type"].Value;
                        List<object> avalue = new List<object>();
                        foreach (var str in sdata.Split(","))
                        {
                            avalue.Add(convert(str, atype));
                        }
                        value = avalue;
                    }
                    else
                    {
                        value = convert(data, type);
                    }
                    item.Add(key, value);
                }
                _data.Add(item);
            }
        }

        /// <summary>
        /// convert       
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private object convert(object data, string type)
        {
            object value;
            switch (type)
            {
                case "int":
                    value = Convert.ToInt64(data);
                    break;
                case "float":
                    value = Convert.ToDouble(data);
                    break;
                case "boolean":
                    value = Convert.ToBoolean(data);
                    break;
                case "date":
                    value = Convert.ToDateTime(data);
                    break;
                case "object":
                    value = JsonConvert.DeserializeObject(data == null ? "{}" : data.ToString());
                    break;
                default:
                    value = data == null ? string.Empty : data;
                    break;
            }
            return value;
        }

        /// <summary>
        /// SaveToFile
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="encoding"></param>
        /// <param name="exportArray"></param>
        /// <param name="dateFormat"></param>
        public void SaveToFile(string filePath, Encoding encoding, bool exportArray, string dateFormat)
        {
            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };
            string context;
            if (exportArray)
            {
                context = JsonConvert.SerializeObject(_data, jsonSettings);
            }
            else
            {
                var obj = _data.Count > 0 ? _data[0] : new Dictionary<string, object>();
                context = JsonConvert.SerializeObject(obj, jsonSettings);
            }
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                {
                    writer.Write(context);
                }
            }
        }
    }
}
