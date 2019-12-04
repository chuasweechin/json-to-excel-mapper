using System.IO;
using System.Configuration;

namespace JsonToExcelMapperTool.Utils
{
  class Config
  {
		internal static readonly string JSON_FILE_EXTENSION = ConfigurationManager.AppSettings["jsonFileExtension"];
		internal static readonly string JSON_FOLDER_DIRECTORY = Path.GetFullPath(ConfigurationManager.AppSettings["jsonFolderDirectory"]);

		internal static readonly string EXCEL_TEMPLATE_DATA_STORE_WORKSHEET_NAME = ConfigurationManager.AppSettings["excelTemplateDataStoreWorksheetName"];
		internal static readonly string EXCEL_TEMPLATE_FULL_PATH = Path.GetFullPath(ConfigurationManager.AppSettings["excelTemplateFullPath"]);
  }
}
