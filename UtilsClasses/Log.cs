using System.Collections.Generic;
using System.IO;

namespace RevitToGOST
{
	static class Log
	{
		public static readonly string LogPath = "...";
		public static void PrintList<T>(List<T> list)
		{
			foreach (var element in list)
			{
				Log.Write(element.ToString());
				Log.Write(" ");
			}
			Log.Write("\n");
		}

		public static void ClearLog()
		{
			File.WriteAllText(LogPath, string.Empty);
		}

		public static void Write(string format, params object[] objs)
		{
			using (StreamWriter sw = File.AppendText(LogPath))
				sw.Write(format, objs);
		}

		public static void WriteLine(string format, params object[] objs)
		{
			using (StreamWriter sw = File.AppendText(LogPath))
				sw.WriteLine(format, objs);
		}
	}
}
