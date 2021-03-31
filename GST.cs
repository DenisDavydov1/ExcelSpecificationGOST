using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace RevToGOSTv0
{
	class GST
	{
		/*
		** Member fields:
		*/

		public string Name { get; set; }
		public string Type { get; set; }
		public string Format { get; set; }
		public string Orientation { get; set; }

		// Positions map:
		//  _____
		// | 1 2 |
		// | 3 4 |
		//  ‾‾‾‾‾
		public int Position { get; set; }
		public List<int[]> Columns;
		public List<int[]> Rows;
		public List<List<int[]>> Fields;
		public string[] HeaderList { get; set; }
		public string Font { get; set; }
		public int FontSize { get; set; }
		public List<int[]> FontSizes { get; set; }

		// Vertical alignment:
		// 4 = Top
		// 1 = Center
		// 0 = Bottom
		public int VerticalAlignment { get; set; }
		public List<int[]> VerticalAlignments { get; set; }

		// Horizontal alignment:
		// 6 = Left
		// 0 = Center
		// 7 = Right
		public int HorizontalAlignment { get; set; }
		public List<int[]> HorizontalAlignments { get; set; }

		public Dictionary<string, int[]> Map { get; set; }
		public bool VerticalText { get; set; }
		public int Line { get; set; }
		public int LinesCount { get; set; }
		public List<List<string>> Data { get; set; }

		/*
		**	Member methods
		*/

		public static GST LoadConfFile(string ConfFilePath)
		{
			if (!File.Exists(ConfFilePath))
				throw new Exception();
			string config = File.ReadAllText(ConfFilePath);
			return JsonConvert.DeserializeObject<GST>(config);
		}

		public void AddData(List<List<string>> data)
		{
			if (Data == null)
				Data = new List<List<string>>(data);
			else
				foreach (List<string> line in data)
					Data.Add(new List<string>(line));
		}
	}
}
