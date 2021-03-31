using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

namespace RevToGOSTv0
{
	class WorkSheet
	{
		/*
		**	Member properties
		*/
		public string Name { get; set; }
		private WorkBook WB;
		private IXLWorksheet WS;
		private string Format;
		private string Orientation;
		private int Height;
		private int Width;
		private List<GST> Tables;
		public List<int[]> Columns;
		public List<int[]> Rows;
		//public List<List<int[]>> Fields;

		/*
		**	Member methods
		*/

		public WorkSheet(WorkBook workbook, string name = Constants.DefaultName)
		{
			Name = name;
			WB = workbook;
			WS = WB.AddWorkSheet(Name);
			Tables = new List<GST>();
		}

		public void AddTable(GST table)
		{
			if (Tables.Count == 0)
			{
				Tables.Add(table);
				Format = table.Format;
				Orientation = table.Orientation;
				Columns = new List<int[]>(table.Columns);
				Rows = new List<int[]>(table.Rows);
				//Fields = new List<List<int[]>>(table.Fields); // XMLTools.CleanFields(table.Fields);
				SetPageDimensions();
			}
			else
			{
				//if (Format != table.Format || Orientation != table.Orientation)
				//	return;
				if (table.Position > 1)
				{
					XMLTools.CompleteFields(table, Height, Width);
					XMLTools.CreateNormalTable(table, Height, Width);
				}
				Tables.Add(table);
				if (table.Columns != null && table.Rows != null)
					(Columns, Rows) = XMLTools.MergeTables(Columns, Rows, table.Columns, table.Rows);
				//if (table.Fields != null)
				//	Fields = XMLTools.MergeFields(Fields, table.Fields); // XMLTools.CleanFields(Fields.Concat(table.Fields).ToList());
			}
		}

		private void SetPageDimensions()
		{
			if (Format == "A3" && Orientation == "Portrait")
				(Height, Width) = (Constants.A3.Height, Constants.A3.Width);
			else if (Format == "A3" && Orientation == "Landscape")
				(Height, Width) = (Constants.A3.Width, Constants.A3.Height);
			else if (Format == "A4" && Orientation == "Portrait")
				(Height, Width) = (Constants.A4.Height, Constants.A4.Width);
			else if (Format == "A4" && Orientation == "Landscape")
				(Height, Width) = (Constants.A4.Width, Constants.A4.Height);
			else if (Format == "A11" && Orientation == "Portrait")
				(Height, Width) = (Constants.A11.Height, Constants.A11.Width);
			else if (Format == "A11" && Orientation == "Landscape")
				(Height, Width) = (Constants.A11.Width, Constants.A11.Height);
		}

		public void BuildWorkSheet()
		{
			ApplyParameters();
			ReshapeWorkSheet();
			AlignBorders();
			MergeCells();
			ApplyStyles();
			FillHeader();

			//Log.WriteLine("Indexes: {0} (10, 20)", XMLTools.GetCellIndexesBySize(Rows, Columns, 19, 19));
			//Log.WriteLine("Indexes: {0} (11, 21)", XMLTools.GetCellIndexesBySize(Rows, Columns, 19, 19));
			//Log.WriteLine("Indexes: {0} (20, 90)", XMLTools.GetCellIndexesBySize(Rows, Columns, 20, 90));
			//Log.WriteLine("");
		}

		private void ApplyParameters()
		{
			//// Page setup
			// Format
			if (Format == "A4")
				WS.PageSetup.PaperSize = XLPaperSize.A4Paper;
			else if (Format == "A3")
				WS.PageSetup.PaperSize = XLPaperSize.A3Paper;
			else
				WS.PageSetup.PaperSize = XLPaperSize.A4Paper;
			// Orientation
			if (Orientation == "Landscape")
				WS.PageSetup.PageOrientation = XLPageOrientation.Landscape;
			else if (Orientation == "Portrait")
				WS.PageSetup.PageOrientation = XLPageOrientation.Portrait;
			else
				WS.PageSetup.PageOrientation = XLPageOrientation.Default;
			// Adjust scale
			WS.PageSetup.AdjustTo(100);
			// Set dpi
			WS.PageSetup.VerticalDpi = 600;
			WS.PageSetup.HorizontalDpi = 600;

			//// Set font
			// ... to do ....
			//WS.Style.Font.FontName = "GOST A";
			//WS.Column(1).Style.Font.FontName = "GOST A";
			//WS.Column(1).Style.Font.FontSize = 7;
			//WS.Column(1).Style.Font.Italic = true;

			//// Margins setup
			WS.PageSetup.Margins.Top = 0;
			WS.PageSetup.Margins.Left = 0;
			WS.PageSetup.Margins.Right = 0;
			WS.PageSetup.Margins.Bottom = 0;
			WS.PageSetup.Margins.Footer = 0;
			WS.PageSetup.Margins.Header = 0;
			WS.PageSetup.CenterHorizontally = false;
			WS.PageSetup.CenterVertically = false;
		}

		private void ReshapeWorkSheet()
		{
			SortedSet<int> cols = XMLTools.GetSortedSet(Columns);
			SortedSet<int> rows = XMLTools.GetSortedSet(Rows);

			// Set widths
			for (int i = 0, size = cols.ElementAt(0); i < cols.Count; i++)
			{
				if (i > 0)
					size = cols.ElementAt(i) - cols.ElementAt(i - 1);
				WS.Column(i + 1).Width = XMLTools.mmToWidth(size);
			}
			// Set heights
			for (int i = 0, size = rows.ElementAt(0); i < rows.Count; i++)
			{
				if (i > 0)
					size = rows.ElementAt(i) - rows.ElementAt(i - 1);
				WS.Row(i + 1).Height = XMLTools.mmToHeight(size); // size * Constants.mm_h;
			}
		}

		private void AlignBorders()
		{
			SortedSet<int> cols = XMLTools.GetSortedSet(Columns);
			SortedSet<int> rows = XMLTools.GetSortedSet(Rows);

			Array cols_arr = cols.ToArray();
			for (int i = 0; i < Columns.Count; i++)
			{
				for (int j = 0; j < Columns[i].Length; j++)
				{
					int j_real = Array.IndexOf(cols_arr, Columns[i].Take(j + 1).Sum()) + 1;
					if (j_real == 0 || Columns[i].Take(j + 1).Sum() >= cols.Max())
						break;
					WS.Cell(i + 1, j_real).Style.Border.RightBorder = XLBorderStyleValues.Thin;
				}
			}
			Array rows_arr = rows.ToArray();
			for (int i = 0; i < Rows.Count; i++)
			{
				for (int j = 0; j < Rows[i].Length; j++)
				{
					int j_real = Array.IndexOf(rows_arr, Rows[i].Take(j + 1).Sum()) + 1;
					if (j_real == 0 || Rows[i].Take(j + 1).Sum() >= rows.Max())
						break;
					WS.Cell(j_real, i + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
				}
			}
		}

		private void MergeCells()
		{
			foreach (GST table in Tables)
			{
				if (table.Fields == null)
					continue;
				foreach (List<int[]> line in table.Fields)
				{
					foreach (int[] field in line)
					{
						(int y1, int x1) = XMLTools.GetCellIndexesBySize(Rows, Columns, field[0], field[1]);
						(int y2, int x2) = XMLTools.GetCellIndexesBySize(Rows, Columns, field[2], field[3]);
						WS.Range(y1, x1, y2, x2).Merge();
						WS.Range(y1, x1, y2, x2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
						//WS.Cell(y1, x1).Style.Alignment.Vertical = (XLAlignmentVerticalValues)table.VerticalAlignment;
						//WS.Cell(y1, x1).Style.Alignment.Vertical = (XLAlignmentVerticalValues)table.HorizontalAlignment;
					}
				}
			}
			//foreach (int[] field in Fields)
			//{
			//	//Log.WriteLine("Merge: {0} {1} - {2} {3}", field[0], field[1], field[2], field[3]);
			//	(int y1, int x1) = XMLTools.GetCellIndexesBySize(Rows, Columns, field[0], field[1]);
			//	(int y2, int x2) = XMLTools.GetCellIndexesBySize(Rows, Columns, field[2], field[3]);
			//	WS.Range(y1, x1, y2, x2).Merge();
			//	WS.Range(y1, x1, y2, x2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
			//}
		}

		private void ApplyStyles()
		{
			foreach (GST table in Tables)
			{
				if (table.Fields == null)
					continue;
				for (int i = 0; i < table.Fields.Count; ++i)
				{
					for (int j = 0; j < table.Fields[i].Count; ++j)
					{
						// Get cell coordinates
						(int y, int x) = XMLTools.GetCellIndexesBySize(Rows, Columns, table.Fields[i][j][0], table.Fields[i][j][1]);
						
						// Set font
						if (table.Font != null)
							WS.Cell(y, x).Style.Font.FontName = table.Font;
						
						// Set font size
						if (table.FontSizes != null && table.FontSizes[i][j] != 0)
							WS.Cell(y, x).Style.Font.FontSize = table.FontSizes[i][j];
						else if (table.FontSize != 0)
							WS.Cell(y, x).Style.Font.FontSize = table.FontSize;
						
						// Set vertical alignments
						if (table.VerticalAlignments != null && 0 <= table.VerticalAlignments[i][j] && table.VerticalAlignments[i][j] <= 4)
							WS.Cell(y, x).Style.Alignment.Vertical = (XLAlignmentVerticalValues)table.VerticalAlignments[i][j];
						else
							WS.Cell(y, x).Style.Alignment.Vertical = (XLAlignmentVerticalValues)table.VerticalAlignment;
						
						// Set horizontal alignments
						if (table.HorizontalAlignments != null && 0 <= table.HorizontalAlignments[i][j] && table.HorizontalAlignments[i][j] <= 7)
							WS.Cell(y, x).Style.Alignment.Horizontal = (XLAlignmentHorizontalValues)table.HorizontalAlignments[i][j];
						else
							WS.Cell(y, x).Style.Alignment.Horizontal = (XLAlignmentHorizontalValues)table.HorizontalAlignment;
						
						// Apply word wrap
						WS.Cell(y, x).Style.Alignment.WrapText = true;

						// Set vertical text direction
						if (table.VerticalText == true)
							WS.Cell(y, x).Style.Alignment.SetTextRotation(90);
						//WS.Cell(y, x).Style.Alignment.SetTopToBottom();

					}
				}
			}
		}

		private void FillHeader()
		{
			foreach (GST table in Tables)
			{
				if (table.Map == null || table.Fields == null)
					continue;
				foreach (KeyValuePair<string, int[]> entry in table.Map)
				{
					if (entry.Key.First() != '_' && entry.Value.Length == 2)
					{
						int y = table.Fields[entry.Value[0]][entry.Value[1]][0];
						int x = table.Fields[entry.Value[0]][entry.Value[1]][1];
						(y, x) = XMLTools.GetCellIndexesBySize(Rows, Columns, y, x);
						WS.Cell(y, x).Value = entry.Key.Replace("_", string.Empty);
					}
				}
				//if (table.HeaderList == null || table.Fields == null)
				//	continue;
				//for (int i = 0; i < table.HeaderList.Length; ++i)
				//{
				//	(int y, int x) = XMLTools.GetCellIndexesBySize(Rows, Columns, table.Fields[0][i][0], table.Fields[0][i][1]);
				//	WS.Cell(y, x).Value = table.HeaderList[i];
				//}
			}
		}

	} // class WorkSheet
} // namespace RevToGOSTv0

//	"HeaderList":
//	[
//		"Инв. № подп.", "Подп. и дата", "Взам. инв. №", "Согласовано"
//	],

// "HeaderList":
// [
// 	"Поз.",
// 	"Наименование и техническая\r\nхарактеристика",
// 	"Тип, марка,\r\nобозначение\r\nдокумента,\r\nопросного листа",
// 	"Код\r\nпродукции",
// 	"Поставщик",
// 	"Ед.\r\nизме-\r\nре-\r\nния",
// 	"Коли-\r\nчест-\r\nво",
// 	"Масса\r\n1 ед.,\r\nкг",
// 	"Приме-\r\nчание"
// ],