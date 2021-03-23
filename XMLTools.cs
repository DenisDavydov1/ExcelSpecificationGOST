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

		public static double mmToHeight(double mm)
		{
			if (mm <= 0.1)
				return 0.0;
			else if (mm <= 14.0)
			// y = -0,0003x3 + 0,0069x2 + 2,797x - 0,0619
				return -0.0003 * Math.Pow(mm, 3) + 0.0069 * Math.Pow(mm, 2) + 2.797 * mm - 0.0619;
			// y = -2E-06x3 + 0,0004x2 + 2,8122x + 0,2867
			return -2e-06 * Math.Pow(mm, 3) + 0.0004 * Math.Pow(mm, 2) + 2.8122 * mm + 0.2867;
		}

		public static double mmToWidth(double mm)
		{
			if (mm <= 0.1)
				return 0.0;
			else if (mm <= 1.0)
				return 0.00001;
			else if (mm <= 5.0)
				// y = -0,01x3 + 0,145x2 - 0,105x + 0,03
				return -0.01 * Math.Pow(mm, 3) + 0.145 * Math.Pow(mm, 2) - 0.105 * mm + 0.03;
			else if (mm <= 20.0)
				// y = 0,0001x3 - 0,0036x2 + 0,5381x - 0,8111
				return 0.0001 * Math.Pow(mm, 3) - 0.0036 * Math.Pow(mm, 2) + 0.5381 * mm - 0.8111;
			// y = -3E-08x3 + 2E-05x2 + 0,5x - 0,6581
			return -3e-08 * Math.Pow(mm, 3) + 2e-05 * Math.Pow(mm, 2) + 0.5 * mm - 0.6581;
		}

		//public static double mmToSize(double mm)
		//{
		//	if (mm <= 0.1)
		//		return 0.0;
		//	else if (mm <= 3)
		//		return mm * (-0.0276 * Math.Pow(mm, 2) + 0.166 * mm + 0.1557);
		//	else if (3.0 < mm && mm < 100.0)
		//		return mm * (-6e-12 * Math.Pow(mm, 6) + 2e-09 * Math.Pow(mm, 5) -3e-07 * Math.Pow(mm, 4) + 2e-05 * Math.Pow(mm, 3) - -0.0007 * Math.Pow(mm, 2) + 0.0125 * mm + 0.3935); //  - 6E-12x6 + 2E-09x5 - 3E-07x4 + 2E-05x3 - 0, 0007x2 + 0, 0125x + 0, 3935
		//	return mm * 0.5;
		//}

	} // class XMLTools
} // namespace RevToGOSTv0
