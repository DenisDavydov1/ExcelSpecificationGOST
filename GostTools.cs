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

namespace RevitToGOST
{
	static class GostTools
	{
		/*
		** Member methods
		*/

		public static ElementSet ElementConvert(IList<Element> list)
		{
			ElementSet elemSet = new ElementSet();
			foreach (Element elem in list)
				elemSet.Insert(elem);
			return elemSet;
		}

	} // class GostTools

	static class Constants
	{
		public const string LogPath = @"F:\CS_CODE\REVIT\PROJECTS\tmp";
		public const string DefaultFilePath = @"F:\CS_CODE\REVIT\PROJECTS\output.xlsx";
		public const double inch = 2.54;
		public const double mm_w = 0.483;
		public const double mm_h = 2.9;
		public const string DefaultName = "Без названия";
		public struct A3
		{
			public const int Height = 420;
			public const int Width = 297;
		}
		public struct A4
		{
			public const int Height = 297;
			public const int Width = 210;
		}
		// to delete::::
		public struct A11
		{
			public const int Height = 100;
			public const int Width = 60;
		}
	} // static class Constants

	static class ConfFile
	{
		public static readonly string[] Paths = {
			String.Empty,								// None,
			@"F:\CS_CODE\REVIT\PROJECTS\RevitToGOST\Templates\GOST_21_110_2013_Table1.json",	// GOST_21_110_2013_Table1
			@"F:\CS_CODE\REVIT\PROJECTS\RevitToGOST\Templates\GOST_21_101_2020_Dop3.json",	// GOST_P_21_101_2020_Dop3
			@"F:\CS_CODE\REVIT\PROJECTS\RevitToGOST\Templates\GOST_21_101_2020_Stamp3.json",	// GOST_P_21_101_2020_Stamp3
			@"F:\CS_CODE\REVIT\PROJECTS\RevitToGOST\Templates\GOST_21_101_2020_Title12.json"	// GOST_P_21_101_2020_Title_12
		};

		public static readonly int[] Lines = {
			0,	// None,
			24,	// GOST_21_110_2013_Table1
			0,	// GOST_P_21_101_2020_Dop3
			0,	// GOST_P_21_101_2020_Stamp3
			0	// GOST_P_21_101_2020_Title_12
		};

		public static readonly Action<ElementContainer>[] FillLine =
		{
			null,
			GOST_21_110_2013.FillLine,
			null,
			null,
			null
		};
	}

	static class ClassExtensions
	{
		public static int TotalLinesCount(this Dictionary<string, ElementCollection> dict)
		{
			int lines = dict.Keys.Count;
			foreach (ElementCollection elemCol in dict.Values)
				lines += elemCol.Count;
			return lines;
		}
	}

} // namespace RevitToGOST
