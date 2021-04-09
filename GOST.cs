using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System.IO;
using Newtonsoft.Json;

namespace RevitToGOST
{
	public class GOST
	{
		/*
		** Member fields:
		*/

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
		//public ElementSet ElemSet { get; set; }
		public ElementCollection ElemCol { get; set; }

		public IGostData GostData { get; set; }

		public int[] Frame { get; set; }
		public List<int[]> Borders { get; set; }
		public Standarts Standart { get; set; }

		public bool IsFull { get { return ElemCol.Count >= ConfFile.Lines[(int)Standart]; } }

		public enum Types
		{
			None,
			Title,
			Table,
			Stamp,
			Dop
		}

		public enum Standarts
		{
			None,
			GOST_21_110_2013_Table1,
			GOST_P_21_101_2020_Dop3,
			GOST_P_21_101_2020_Stamp3,
			GOST_P_21_101_2020_Title_12
		}

		/*
		**	Member methods
		*/

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

		//public void ApplyGostData()
		//{
		//	if (Standart == Standarts.GOST_21_110_2013_Table1)
		//	{
		//		GostData = new GOST_21_110_2013(ElemCol);
		//		GostData.FillLines();	
		//		Data = GostData.FillList();
		//	}
		//	if (Standart == Standarts.GOST_P_21_101_2020_Title_12)
		//	{
		//		GostData = new GOST_P_21_101_2020_Title_12();
		//		GostData.FillLines();
		//		Data = GostData.FillList();
		//	}
		//}

	} // class GOST

} // namespace RevitToGOST
