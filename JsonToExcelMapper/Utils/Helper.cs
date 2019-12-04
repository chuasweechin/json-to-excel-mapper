using System;
using System.Data;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonToExcelMapperTool.Utils
{
  class Helper
  {
		internal static DataTable CreateDataTableWithColumnHeader(params string[] headerValues)
		{
			DataTable dataTable = new DataTable();

			for (int i = 0; i < headerValues.Length; i++)
			{
				dataTable.Columns.Add(headerValues[i]);
			}

			return dataTable;
		}

		internal static void SetDataTableRowValue(ref DataTable dataTable, params string[] values)
		{
			DataRow row = dataTable.NewRow();

			for (int i = 0; i < values.Length; i++)
			{
				row[i] = values[i];
			}

			dataTable.Rows.Add(row);
		}

		internal static void WriteDataTableToWorksheet(ref Excel.Worksheet xlsWorksheet, DataTable dataTable)
		{
			int rowCount = dataTable.Rows.Count;
			int colCount = dataTable.Columns.Count;

			object[,] _arr = new object[rowCount, colCount];

			for (int row = 0; row < rowCount; row++)
			{
				for (int col = 0; col < colCount; col++)
				{
					_arr[row, col] = dataTable.Rows[row][col];
				}
			}

			Excel.Range start = (Excel.Range)xlsWorksheet.Cells[2, 1];
			Excel.Range end = (Excel.Range)xlsWorksheet.Cells[rowCount + 1, colCount];

			Excel.Range range = xlsWorksheet.get_Range(start, end);
			range.Value = _arr;
		}

		internal static void ApplyExistingFilterOnWorksheets(ref Excel.Workbook xlsWorkBook)
		{
			foreach (Excel.Worksheet ws in xlsWorkBook.Worksheets)
			{
				if (ws.AutoFilter != null)
				{
					ws.AutoFilter.ApplyFilter();
				}
			}
		}

		internal static void WriteJsonObjectToDataTable(ref DataTable dataTable, JObject jsonObject)
		{
			if (jsonObject["data"] != null)
			{
				foreach (JToken token in jsonObject["data"].Children())
				{
					ReadJsonObjectPathAndValue(token, jsonObject, ref dataTable);
				}
			}
			else
			{
				throw new FormatException();
			}
		}

		internal static void ReadJsonObjectPathAndValue(JToken token, JObject JsonObject, ref DataTable dataTable)
		{
			if (JsonObject.SelectToken(token.Path).GetType().ToString() == "Newtonsoft.Json.Linq.JValue")
			{
				SetDataTableRowValue(ref dataTable, (dataTable.Rows.Count + 1).ToString(), token.Path, JsonObject.SelectToken(token.Path).ToString());
				return;
			}

			foreach (JToken obj in token.Values())
			{
				if (obj.Type.ToString() == "Object")
				{
					foreach (JToken innerObj in obj.Values())
					{
					ReadJsonObjectPathAndValue(innerObj, JsonObject, ref dataTable);
					}
				}
				else
				{
					ReadJsonObjectPathAndValue(obj, JsonObject, ref dataTable);
				}
			}
		}
  }
}


















