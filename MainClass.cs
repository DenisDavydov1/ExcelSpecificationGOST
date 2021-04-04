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
	[TransactionAttribute(TransactionMode.Manual)]
	[RegenerationAttribute(RegenerationOption.Manual)]
	public class MainClass : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			Log.ClearLog();

			Rvt.Handler = new RvtHandler(commandData, elements);
			Rvt.Data = new RvtData();
			Rvt.Control = new RvtControl();

			MainWindow mainWin = new MainWindow();
			mainWin.Show();

			return Result.Succeeded;
		}
	} // class MainClass

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

} // namespace RevitToGOST
