using System;

namespace JsonToExcelMapperTool.Utils
{
  class Logger
  {
		internal static void Info(string message)
		{
			Console.WriteLine($"[Info] { message }");
		}

		internal static void Error(string message)
		{
			Console.WriteLine($"[Program Crashed] { message }");
		}
  }
}
