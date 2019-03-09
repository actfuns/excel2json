using ExcelDataReader;
using System;
using System.Data;
using System.IO;

namespace ActFuns.Tools.Excel2Json
{
    public class ExcelLoader
    {
        /// <summary>
        /// data
        /// </summary>
        private DataSet mData;

        /// <summary>
        /// Sheets
        /// </summary>
        public DataTableCollection Sheets
        {
            get
            {
                return mData.Tables;
            }
        }

        /// <summary>
        /// ExcelLoader
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelLoader(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    mData = result;
                }
            }

            if (Sheets.Count < 1)
            {
                throw new Exception("Excel file is empty: " + filePath);
            }
        }
    }
}
