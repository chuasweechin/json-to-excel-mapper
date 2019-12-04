using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Reflection;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using JsonToExcelMapperTool.Utils;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace JsonToExcelMapperTool
{
  class Program
  {
		static void Main(string[] args)
		{
			List<JObject> jsonObjects = ReadJsonFiles(
				Config.JSON_FILE_EXTENSION,
				Config.JSON_FOLDER_DIRECTORY
			);

			WriteObjectsToExcel(
				jsonObjects,
				Config.JSON_FOLDER_DIRECTORY,
				Config.EXCEL_TEMPLATE_DATA_STORE_WORKSHEET_NAME,
				Config.EXCEL_TEMPLATE_FULL_PATH
			);
		}

		private static List<JObject> ReadJsonFiles(string jsonFileExtension, string jsonFolderDirectory)
		{
			try
			{
				Logger.Info($"Searching for json files in, { jsonFolderDirectory }......");

				List<JObject> result = new List<JObject>();
				List<string> filesInDirectory = Directory.EnumerateFiles(jsonFolderDirectory, $"*.{ jsonFileExtension }").ToList();

				if (filesInDirectory.Count() == 0)
				{
					throw new FileNotFoundException();
				}

				foreach (string file in filesInDirectory)
				{
					Logger.Info($"Json file detected: { Path.GetFileName(file) }");

					using (StreamReader r = new StreamReader(Path.GetFullPath(file)))
					{
						JObject jsonObj = JObject.Parse(r.ReadToEnd());
						jsonObj.Add("Filename", Path.GetFileNameWithoutExtension(file));
						result.Add(jsonObj);
					}
				}

				Logger.Info("Complete json file(s) search.");
				return result;

				}
				catch (FileNotFoundException)
				{
					Logger.Error($"Json file not found exception: No json file found in the directory { jsonFolderDirectory }. Please check your directory.");
				}
				catch (DirectoryNotFoundException ex)
				{
					Logger.Error($"Invalid directory exception: { ex.Message }");
				}
				catch (ArgumentException ex)
				{
					Logger.Error($"Malformed json format exception: { ex.Message }");
				}
				catch (Exception ex)
				{
					if (ex.InnerException is FormatException)
					{
						Logger.Error($"Json data format exception: { ex.Message }");
					}
					else
					{
						Logger.Error($"Unhandle json file read exception: { ex.Message }");
					}
			}
			return null;
		}

		private static void WriteObjectsToExcel(List<JObject> jsonObjects, string jsonFolderDirectory, string excelTemplateDataStoreWorksheetName, string excelTemplateFile)
		{
			Excel.Application app = new Excel.Application();

			if (app == null)
			{
				Logger.Error($"Excel is missing from your system. Please install excel to your system before using the program.");
				return;
			}

			if (jsonObjects == null || jsonObjects.Count == 0)
			{
				Logger.Error($"There is an error reading your json files. Please check your json files.");
				return;
			}

			try
			{
				Logger.Info($"Start writing json content to Excel....");

				foreach (JObject jsonObject in jsonObjects)
				{
					DataTable dataTable = Helper.CreateDataTableWithColumnHeader("S/N", "Data_ID", "Value");
					string saveAsFilename = Formatter.FormatOutputFile(jsonFolderDirectory, jsonObject["Filename"].ToString(), ".xlsx");

					Excel.Workbook xlsWorkBook = app.Workbooks.Open(@excelTemplateFile);
					Excel.Worksheet xlsDataStoreWorksheet = (Excel.Worksheet)xlsWorkBook.Worksheets.Item[excelTemplateDataStoreWorksheetName];

					Helper.WriteJsonObjectToDataTable(
						ref dataTable,
						jsonObject
					);

					Helper.WriteDataTableToWorksheet(
						ref xlsDataStoreWorksheet,
						dataTable
					);

					Helper.ApplyExistingFilterOnWorksheets(
						ref xlsWorkBook
					);

					xlsWorkBook.SaveAs(
						saveAsFilename, Excel.XlFileFormat.xlOpenXMLWorkbook,
						Missing.Value, Missing.Value, Missing.Value, Missing.Value,
						Excel.XlSaveAsAccessMode.xlExclusive,
						Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value
					);

					xlsWorkBook.Close(0);
					Marshal.ReleaseComObject(xlsWorkBook);
				}

				Logger.Info("Complete writing json content to Excel.");
			}
			catch (FormatException)
			{
				Logger.Error($"Json object is missing [Data] as a root field. Please move your payload into [Data] field in the json file.");
			}
			catch (PathTooLongException e)
			{
				Logger.Error($"Output file path too long exception: { e.Message }. Please shorter the filename of the json file and try again.");
			}
			catch (COMException e)
			{
				Logger.Error($"{ e.Message }. You might be already opened the excel that the program is trying to write to. Please ensure that the excel file is closed. If after closing the excel file and it still does not work, try restarting your system in case hidden processes is locking the file.");
			}
			catch (Exception ex)
			{
				Logger.Error($"Unhandle exception: { ex.Message }");
			}
			finally
			{
				app.Quit();
				Marshal.ReleaseComObject(app);
			}
		}
	}
}