using System;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace ExcelToJsonLib
{
    /// <summary>
    /// 将 Excel 文件(*.xls 或者 *.xlsx)加载到内存 DataSet
    /// </summary>
    public class ExcelLoader
    {
        private DataSet data;
        public ExcelLoader(string filePath, int headerRow)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(CreateDataSetReadConfig(headerRow));
                    data = result;
                }
            }

            if (Sheets.Count < 1)
            {
                throw new Exception("Excel file is empty: " + filePath);
            }
        }

        public DataTableCollection Sheets
        {
            get
            {
                return this.data.Tables;
            }
        }

        private ExcelDataSetConfiguration CreateDataSetReadConfig(int headerRow)
        {
            var tableConfig = new ExcelDataTableConfiguration()
            {
                UseHeaderRow = true,
                FilterRow = (rowReader) =>
                {
                    return rowReader.Depth > headerRow - 1;
                },
            };

            return new ExcelDataSetConfiguration()
            {
                UseColumnDataType = true,
                ConfigureDataTable = (tableReader) => { return tableConfig; },
            };
        }
    }
}