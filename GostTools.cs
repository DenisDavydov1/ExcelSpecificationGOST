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
using System.Windows.Media.Imaging;

using RevitToGOST.Properties;
using System.Drawing;
using System.Windows.Data;
using System.IO;

namespace RevitToGOST
{
	static class GostTools
	{
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

	} // static class Constants

	static class ConfFile
	{
		public static readonly string[] Conf = {
			String.Empty,
			Encoding.UTF8.GetString(Resources.GOST_21_110_2013_Table_1),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Stamp_3),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Dop_3),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Stamp_4),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Dop_4),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Stamp_5),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Dop_5),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Stamp_6),
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Dop_6),
			String.Empty,	// table 7
			String.Empty,	// table 8
			String.Empty,	// misc 9
			String.Empty,	// misc 9a
			String.Empty,	// misc 10
			String.Empty,	// misc 11
			Encoding.UTF8.GetString(Resources.GOST_21_101_2020_Title_12),
			String.Empty,	// misc 13
			String.Empty	// title 14
		};

		public static readonly int[] Lines = {
			0,	// None
			24,	// GOST_21_110_2013_Table1
			0,	// GOST_P_21_101_2020_Stamp3
			0,	// GOST_P_21_101_2020_Dop3
			0,	// stamp 4
			0,	// dop 4
			0,	// stamp 5
			0,	// dop 5
			0,	// stamp 6
			0,	// dop 6
			0,	// table 7
			0,	// table 8
			0,	// misc 9
			0,	// misc 9a
			0,	// misc 10
			0,	// misc 11
			0,	// GOST_P_21_101_2020_Title_12
			0,	// misc 13
			0	// title 14
		};

		public static readonly Action<int>[] FillTitle = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			null,	// GOST_P_21_101_2020_Stamp3
			null,	// GOST_P_21_101_2020_Dop3
			null,	// stamp 4
			null,	// dop 4
			null,	// stamp 5
			null,	// dop 5
			null,	// stamp 6
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			GOST_P_21_101_2020_Title_12.FillTitle,
			null,	// misc 13
			null	// title 14
		};

		public static readonly Action<ElementContainer>[] FillLine = {
			null,	// None
			GOST_21_110_2013.FillLine,
			null,	// GOST_P_21_101_2020_Stamp3
			null,	// GOST_P_21_101_2020_Dop3
			null,	// stamp 4
			null,	// dop 4
			null,	// stamp 5
			null,	// dop 5
			null,	// stamp 6
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// misc 13
			null	// title 14
		};

		public static readonly Action<int>[] FillStamp = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			GOST_P_21_101_2020_Stamp_3.FillStamp,
			null,	// GOST_P_21_101_2020_Dop3
			GOST_P_21_101_2020_Stamp_4.FillStamp,
			null,	// dop 4
			GOST_P_21_101_2020_Stamp_5.FillStamp,
			null,	// dop 5
			GOST_P_21_101_2020_Stamp_6.FillStamp,
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// misc 13
			null	// title 14
		};

		public static readonly Action<int>[] FillDop = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			null,	// GOST_P_21_101_2020_Stamp3
			GOST_P_21_101_2020_Dop_3.FillDop,
			null,	// stamp 4
			GOST_P_21_101_2020_Dop_4.FillDop,
			null,	// stamp 5
			GOST_P_21_101_2020_Dop_5.FillDop,
			null,	// stamp 6
			GOST_P_21_101_2020_Dop_6.FillDop,
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// misc 13
			null	// title 14
		};

		/*
			// None
			// GOST_21_110_2013_Table1
			// GOST_P_21_101_2020_Stamp3
			// GOST_P_21_101_2020_Dop3
			// stamp 4
			// dop 4
			// stamp 5
			// dop 5
			// stamp 6
			// dop 6
			// table 7
			// table 8
			// misc 9
			// misc 9a
			// misc 10
			// misc 11
			// GOST_P_21_101_2020_Title_12
			// misc 13
			// title 14 
		*/
	} // class ConfFile

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
