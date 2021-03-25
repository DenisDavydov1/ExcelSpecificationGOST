using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;


/*
	"Fields":
	[
		[ 20, 20 ], [ 30, 20 ], [ 40, 20 ]
	]
*/
namespace RevToGOSTv0
{
	class GST
	{
		public string Name { get; set; }
		public string Type { get; set; }
		public string Format { get; set; }
		public string Orientation { get; set; }
		//public int[][] Columns;
		//public int[][] Rows;
		public List<int[]> Columns;
		public List<int[]> Rows;
		public List<List<int[]>> Fields;
		public string[] HeaderList { get; set; }

		public static GST LoadConfFile(string ConfFilePath)
		{
			if (!File.Exists(ConfFilePath))
				throw new Exception();
			string config = File.ReadAllText(ConfFilePath);
			return JsonConvert.DeserializeObject<GST>(config);
		}
	}
}
