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
			else if (mm <= 5.0)
				// y = 0,0891x4 - 1,0159x3 + 3,8964x2 - 3,0026x + 3,121
				return 0.0891 * Math.Pow(mm, 4) - 1.0159 * Math.Pow(mm, 3) + 3.8964 * Math.Pow(mm, 2) - 3.0026 * mm + 3.121;
			else if (mm <= 10.0)
				// y = 0,0907x4 - 2,8895x3 + 34,152x2 - 174,73x + 343,47
				return 0.0907 * Math.Pow(mm, 4) - 2.8895 * Math.Pow(mm, 3) + 34.152 * Math.Pow(mm, 2) - 174.73 * mm + 343.47;
			//else if (mm <= 15.0)
			//	// y = 0,02x4 - 1,23x3 + 27,675x2 - 268,47x + 980,99
			//	return 0.02 * Math.Pow(mm, 4) - 1.23 * Math.Pow(mm, 3) + 27.675 * Math.Pow(mm, 2) - 268.47 * mm + 980.99;
			//// y = 7E-07x4 - 0,0002x3 + 0,0128x2 + 2,4019x + 5,0928
			//return 7e-07 * Math.Pow(mm, 4) - 0.0002 * Math.Pow(mm, 3) + 0.0128 * Math.Pow(mm, 2) + 2.4019 * mm + 5.0928;
			return mm * Constants.mm_h;
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

		private static (SortedSet<int>, SortedSet<int>) ExtractSizes(List<int[]> col, List<int[]> row)
		{
			SortedSet<int> col_extr = new SortedSet<int>();
			SortedSet<int> rows_extr = new SortedSet<int>();

			for (int i = 0; i < col.Count; i++)
				for (int j = 0; j < col[i].Length; j++)
					col_extr.Add(col[i].Take(j + 1).Sum());
			col_extr.Remove(0);

			for (int i = 0; i < row.Count; i++)
				for (int j = 0; j < row[i].Length; j++)
					rows_extr.Add(row[i].Take(j + 1).Sum());
			rows_extr.Remove(0);

			return (col_extr, rows_extr);
		}

		public static SortedSet<int> GetSet(int[] arr)
		{
			SortedSet<int> output = new SortedSet<int>();
			output.Add(arr.Take(1).Sum());
			for (int i = 1; i < arr.Length; ++i)
				output.Add(arr.Take(i + 1).Sum());
			return output;
		}

		private static int GetNearest(ICollection<int> set, int num)
		{
			foreach (int item in set)
				if (item >= num)
					return item;
			return num;
		}

		public static (List<int[]>, List<int[]>) MergeTables(
			List<int[]> col1, List<int[]> row1, List<int[]> col2, List<int[]> row2)
		{
			List<int[]> col = new List<int[]>();
			List<int[]> row = new List<int[]>();

			(SortedSet<int> ecol1, SortedSet<int> erow1) = ExtractSizes(col1, row1);
			(SortedSet<int> ecol2, SortedSet<int> erow2) = ExtractSizes(col2, row2);
			(SortedSet<int> ecol, SortedSet<int> erow) = (new SortedSet<int>(ecol1), new SortedSet<int>(erow1));
			ecol.UnionWith(ecol2);
			erow.UnionWith(erow2);
			//Console.WriteLine("ecol1: " + String.Join(", ", ecol1));
			//Console.WriteLine("erow1: " + String.Join(", ", erow1));
			//Console.WriteLine("ecol2: " + String.Join(", ", ecol2));
			//Console.WriteLine("erow2: " + String.Join(", ", erow2));
			//Console.WriteLine("ecol: " + String.Join(", ", ecol));
			//Console.WriteLine("erow: " + String.Join(", ", erow));

			for (int i = 0; i < erow.Count; ++i)
			{
				int[] c1 = col1[Array.IndexOf(erow1.ToArray(), GetNearest(erow1, erow.ElementAt(i)))];
				int[] c2 = col2[Array.IndexOf(erow2.ToArray(), GetNearest(erow2, erow.ElementAt(i)))];
				(SortedSet<int> l1, SortedSet<int> l2) = (GetSet(c1), GetSet(c2));
				l1.UnionWith(l2);
				col.Add(new int[l1.Count]);
				for (int j = 0; j < l1.Count; ++j)
				{
					if (j == 0)
						col.Last()[j] = l1.ElementAt(j);
					else
						col.Last()[j] = l1.ElementAt(j) - l1.ElementAt(j - 1);
				}
			}

			for (int i = 0; i < ecol.Count; ++i)
			{
				int ind1 = Array.IndexOf(ecol1.ToArray(), GetNearest(ecol1, ecol.ElementAt(i)));
				int ind2 = Array.IndexOf(ecol2.ToArray(), GetNearest(ecol2, ecol.ElementAt(i)));
				int[] r1 = row1[ind1];
				int[] r2 = row2[ind2];
				(SortedSet<int> l1, SortedSet<int> l2) = (GetSet(r1), GetSet(r2));
				l1.UnionWith(l2);
				row.Add(new int[l1.Count]);
				for (int j = 0; j < l1.Count; ++j)
				{
					if (j == 0)
						row.Last()[j] = l1.ElementAt(j);
					else
						row.Last()[j] = l1.ElementAt(j) - l1.ElementAt(j - 1);
				}
			}
			//foreach (var i in col)
			//	Console.WriteLine(String.Join(",", i));
			//Console.WriteLine();
			//foreach (var i in row)
			//	Console.WriteLine(String.Join(",", i));
			return (col, row);
		}

	} // class XMLTools
} // namespace RevToGOSTv0
