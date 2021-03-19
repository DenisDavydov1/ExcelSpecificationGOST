using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace RevToGOSTv0
{
	static class XMLTools
	{
		private const string ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		public static Tuple<int, int> ToNumericCoordinates(string coordinates)
		{
			string first = string.Empty;
			string second = string.Empty;

			CharEnumerator ce = coordinates.GetEnumerator();
			while (ce.MoveNext())
				if (char.IsLetter(ce.Current))
					first += ce.Current;
				else
					second += ce.Current;

			int i = 0;
			ce = first.GetEnumerator();
			while (ce.MoveNext())
				i = (26 * i) + ALPHABET.IndexOf(ce.Current) + 1;

			string str = i.ToString();
			return Tuple.Create(Convert.ToInt32(second), Convert.ToInt32(str));
		}

		//public static IXLWorksheet CreateFormattedWorksheetFurniture(IXLWorkbook workbook)
		//{
		//	IXLWorksheet ws = workbook.Worksheets.Add("Furniture");
		//	this.SetFormat_A3_landscape(ref ws);
		//	this.FormatHeader_GOST_21_110_2013(ref ws);
		//	//ws.Row(1).Height = 20;
		//	//ws.Column(1).Width = 50;
		//	//ws.Column(1).Style.Font.FontName = "GOST A";
		//	//ws.Column(1).Style.Font.FontSize = 14;
		//	//ws.Column(1).Style.Font.Italic = true;

		//	//ws.Cells("5:6").Style.Fill.BackgroundColor = XLColor.Raspberry;
		//	//worksheet.Cell(1, 1).Comment.Style.Alignment.SetAutomaticSize();
		//	//FillCells(ws, "Наименование фурнитуры", XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center,
		//	//	"GOST A", "bold italic", 16, "A1:A1", XLBorderStyleValues.Medium);

		//	return ws;
		//}

		public static void FillCells(
			IXLWorksheet worksheet,
			string text,
			XLAlignmentHorizontalValues horizontalAlignment,
			XLAlignmentVerticalValues verticalAlignment,
			string font,
			string fontStyle,
			int fontSize,
			string range,
			XLBorderStyleValues style)
		{
			string first = range.Substring(0, range.IndexOf(':'));
			string second = range.Substring(first.Length + 1);
			var a = ToNumericCoordinates(first);
			var b = ToNumericCoordinates(second);
			string cell = a.Item1 < b.Item1 ? first : second;

			worksheet.Cell(cell).Value = text;
			worksheet.Cell(cell).Style.Font.FontName = font;
			if (fontStyle.Contains("talic"))
				worksheet.Cell(cell).Style.Font.Italic = true;
			if (fontStyle.Contains("old"))
				worksheet.Cell(cell).Style.Font.Bold = true;
			worksheet.Cell(cell).Style.Font.FontSize = fontSize;
			worksheet.Cell(cell).Style.Alignment.Horizontal = horizontalAlignment;
			worksheet.Cell(cell).Style.Alignment.Vertical = verticalAlignment;
			if (cell == first)
				worksheet.Range(range).Merge();
			else
				worksheet.Range(string.Format("{0}:{1}", second, first)).Merge();
			SetCellsBorders(worksheet, range, style);
		}

		public static void SetCellsBorders(IXLWorksheet worksheet, string range, XLBorderStyleValues style)
		{
			string first = range.Substring(0, range.IndexOf(':'));
			string second = range.Substring(first.Length + 1);
			var a = ToNumericCoordinates(first);
			var b = ToNumericCoordinates(second);

			for (int i = Math.Min(a.Item2, b.Item2); i <= Math.Max(a.Item2, b.Item2); i++)
			{
				worksheet.Cell(Math.Min(a.Item1, b.Item1), i).Style.Border.TopBorder = style;
				worksheet.Cell(Math.Max(a.Item1, b.Item1), i).Style.Border.BottomBorder = style;
			}
			for (int i = Math.Min(a.Item1, b.Item1); i <= Math.Max(a.Item1, b.Item1); i++)
			{
				worksheet.Cell(i, Math.Min(a.Item2, b.Item2)).Style.Border.LeftBorder = style;
				worksheet.Cell(i, Math.Max(a.Item2, b.Item2)).Style.Border.RightBorder = style;
			}
		}


	} // class XMLTools
} // namespace RevToGOSTv0
