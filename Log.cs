using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevToGOSTv0
{
	static class Log
	{

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
			File.WriteAllText(Constants.LogPath, string.Empty);
		}

		public static void Write(string text = "")
		{
			using (StreamWriter sw = File.AppendText(Constants.LogPath))
				sw.Write(text);
		}

		public static void WriteLine(string text = "")
		{
			using (StreamWriter sw = File.AppendText(Constants.LogPath))
				sw.WriteLine(text);
		}
	} // class Log

} // namespace RevToGOSTv0
