using System;
using System.IO;

namespace JsonToExcelMapperTool.Utils
{
  class Formatter
  {
		internal static string FormatOutputFile(string jsonFolderDirectory, string filename, string fileExtension)
		{
			string timeStamp = DateTime.Now.ToString(".ddMMyyyy.HHmmss");
			string outputFilename = $"{ filename }{ timeStamp }{ fileExtension }";
			string fullPathOutputFilename = Path.Combine(jsonFolderDirectory, outputFilename);

			if (fullPathOutputFilename.Length > 255)
			{
				throw new PathTooLongException();
			}

			return fullPathOutputFilename;
		}
  }
}
