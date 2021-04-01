﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

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
		public ElementSet ElemSet { get; set; }

		public IGostData GostData { get; set; }

		public int[] Frame { get; set; }
		public List<int[]> Borders { get; set; }
		public int Form { get; set; }

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

		public void AddData(List<string> line)
		{
			if (Data == null)
				Data = new List<List<string>>() { new List<string>(line) };
			else
				Data.Add(new List<string>(line));
		}

		public void AddElement(ElementSet elementSet)
		{
			if (ElemSet == null)
				ElemSet = new ElementSet();
			ElementSetIterator it = elementSet.ForwardIterator();
			while (it.MoveNext())
				ElemSet.Insert((Element)it.Current);
		}

		public void AddElement(Element element)
		{
			if (ElemSet == null)
				ElemSet = new ElementSet();
			ElemSet.Insert(element);
		}

		public void ApplyGostData()
		{
			if (Name == "ГОСТ 21.110—2013" && Type == "Page")
			{
				GostData = new GOST_21_110_2013(ElemSet);
				GostData.FillLines();
				Data = GostData.FillList();
			}
			if (Name == "ГОСТ 21.101—2020" && Form == 12)
			{
				GostData = new GOST_P_21_101_2020_Title_12();
				GostData.FillLines();
				Data = GostData.FillList();
			}
		}
	}
}
