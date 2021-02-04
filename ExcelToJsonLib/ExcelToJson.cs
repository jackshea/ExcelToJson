using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;

namespace ExcelToJsonLib
{
    public class ExcelToJson
    {
        public JsonSerializerSettings JsonSettings { get; set; }

        public Dictionary<string, Type> TypeMapper = new Dictionary<string, Type>();

        public ExcelToJson()
        {
            JsonSettings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented
            };
            InitTypeMapper();
        }

        private void InitTypeMapper()
        {
            TypeMapper.Add("int", typeof(int));
            TypeMapper.Add("float", typeof(float));
            TypeMapper.Add("double", typeof(double));
            TypeMapper.Add("string", typeof(string));
            TypeMapper.Add("object", typeof(object));
            //TypeMapper.Add("bool", typeof(bool));
            TypeMapper.Add("int[]", typeof(int[]));
            TypeMapper.Add("float[]", typeof(float[]));
            TypeMapper.Add("string[]", typeof(string[]));
            TypeMapper.Add("bool[]", typeof(bool[]));
            TypeMapper.Add("object[]", typeof(object[]));
        }

        public string OpenExcelAndToJson(string filePath)
        {
            var excelLoader = new ExcelLoader(filePath, 1);
            return DataTableToJson(excelLoader.Sheets[0]);
        }

        public string DataTableToJson(DataTable dt)
        {
            if (dt.Rows.Count <= 2)
            {
                throw new Exception("No data.");
            }

            DataRow typeRow = dt.Rows[0];
            Type[] types = new Type[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                types[i] = ConvertToType(typeRow[i].ToString());
            }

            var data = new List<Dictionary<string, object>>();
            for (int i = 2; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                Dictionary<string, object> rowData = new Dictionary<string, object>();
                for (var j = 0; j < dt.Columns.Count; j++)
                {
                    DataColumn dc = dt.Columns[j];
                    var cell = dr[dc];
                    if (types[j] == typeof(string))
                    {
                        rowData[dc.ColumnName] = cell;
                    }
                    else if (types[j] == typeof(int))
                    {
                        rowData[dc.ColumnName] = int.Parse(cell.ToString());
                    }
                    else if (types[j] == typeof(float))
                    {
                        rowData[dc.ColumnName] = float.Parse(cell.ToString());
                    }
                    else if (types[j] == typeof(double))
                    {
                        rowData[dc.ColumnName] = double.Parse(cell.ToString());
                    }
                    else if (types[j] == typeof(bool))
                    {
                        rowData[dc.ColumnName] = bool.Parse(cell.ToString());
                    }
                    else
                    {
                        rowData[dc.ColumnName] = JsonConvert.DeserializeObject(cell.ToString(), types[j]);
                    }
                }

                data.Add(rowData);
            }

            return JsonConvert.SerializeObject(data, JsonSettings);
        }

        private Type ConvertToType(string excelTypeName)
        {
            if (TypeMapper.TryGetValue(excelTypeName, out var type))
            {
                return type;
            }
            return typeof(string);
        }
    }
}
