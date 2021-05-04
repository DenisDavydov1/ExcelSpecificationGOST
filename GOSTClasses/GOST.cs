using System.Collections.Generic;
using Newtonsoft.Json;

namespace RevitToGOST
{
	public class GOST
	{
		#region properties

		public string Name { get; set; }
		public Types Type { get; set; }
		public string Format { get; set; }
		public string Orientation { get; set; }

		// Positions map:
		//  _____
		// | 1 2 |
		// | 3 4 |
		//  ‾‾‾‾‾
		public int Position { get; set; }
		public List<int[]> Columns { get; set; }
		public List<int[]> Rows { get; set; }
		public List<List<int[]>> Fields { get; set; }
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
		public ElementCollection ElemCol { get; set; }
		public int[] Frame { get; set; }
		public List<int[]> Borders { get; set; }
		public Standarts Standart { get; set; }
		public bool IsFull { get { return ElemCol.Count >= ConfFile.Lines[(int)Standart]; } }

		#endregion properties

		#region enums

		public enum Types
		{
			None,
			Title,
			Table,
			Stamp,
			Dop,
			Misc
		}

		public enum Standarts
		{
			None,
			GOST_21_110_2013_Table_1,
			GOST_P_21_101_2020_Stamp_3,
			GOST_P_21_101_2020_Dop_3,
			GOST_P_21_101_2020_Stamp_4,
			GOST_P_21_101_2020_Dop_4,
			GOST_P_21_101_2020_Stamp_5,
			GOST_P_21_101_2020_Dop_5,
			GOST_P_21_101_2020_Stamp_6,
			GOST_P_21_101_2020_Dop_6,
			GOST_P_21_101_2020_Table_7,
			GOST_P_21_101_2020_Table_8,
			misc9,
			misc9a,
			misc10,
			misc11,
			GOST_P_21_101_2020_Title_12,
			GOST_P_21_101_2020_Title_12a,
			title14,
			GOST_2_104_2006_Stamp_1,
			GOST_2_104_2006_Dop_1,
			GOST_2_104_2006_Stamp_2,
			GOST_2_104_2006_Dop_2,
			GOST_2_104_2006_Stamp_2a,
			GOST_2_104_2006_Dop_2a,
			GOST_21_301_2014_Title_2,
			GOST_P_2_106_2019_Table_1,
			GOST_P_2_106_2019_Table_5
		}

		#endregion enums

		#region methods

		public static GOST LoadConfFile(string config)
		{
			GOST newGOST = JsonConvert.DeserializeObject<GOST>(config);
			newGOST.Data = new List<List<string>>();
			newGOST.ElemCol = new ElementCollection();
			return newGOST;
		}

		public void AddData(List<List<string>> data)
		{
			if (Data == null)
				Data = new List<List<string>>(data);
			else
				foreach (List<string> line in data)
					Data.Add(new List<string>(line));
		}

		public void AddData(List<string> line)
		{
			if (Data == null)
				Data = new List<List<string>>() { new List<string>(line) };
			else
				Data.Add(new List<string>(line));
		}


		public void AddElement(ElementCollection elemCol)
		{
			if (ElemCol == null)
				ElemCol = new ElementCollection();
			foreach (ElementContainer elem in elemCol)
				ElemCol.Add(elem);
		}

		public void AddElement(ElementCollection elemCol, int from, int to) // [from, to)
		{
			if (ElemCol == null)
				ElemCol = new ElementCollection();
			for (int i = from; i < to; ++i)
				ElemCol.Add(elemCol[i]);
		}

		public void AddElement(ElementContainer element)
		{
			if (ElemCol == null)
				ElemCol = new ElementCollection();
			ElemCol.Add(element);
		}

		public void ElemColToData()
		{
			if (Type == Types.Table)
			{
				Data = new List<List<string>>();
				foreach (ElementContainer elemCont in ElemCol)
				{
					Data.Add(elemCont.Line);
				}
			}
		}

		#endregion methods
	}
}
