using System;
using System.IO;
using System.Runtime.InteropServices;
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

namespace RevToGOSTv0
{
	[TransactionAttribute(TransactionMode.Manual)]
	[RegenerationAttribute(RegenerationOption.Manual)]
	public class MainClass : IExternalCommand
	{
		Result IExternalCommand.Execute(
			ExternalCommandData commandData,
			ref string message,
			ElementSet elements)
		{
			Log.ClearLog();

			//UIApplication uiApp = commandData.Application;
			//Document doc = uiApp.ActiveUIDocument.Document;


			Rvt.Handler = new RvtHandler(commandData, elements);
			NewTest();
			//rvt.LogCategory(BuiltInCategory.OST_Furniture);
			//rvt.LogAllElements();
			//rvt.LogCategory(BuiltInCategory.OST_Walls);


			//ProjectInfo pi = doc.ProjectInformation;
			//Log.WriteLine("Author: " + pi.Author);
			//Log.WriteLine("OrganizationName: " + pi.OrganizationName);

			//uiWindow uiWin = new uiWindow(list);
			//uiWin.Show();

			return Result.Succeeded;
		}

		internal void NewTest()
		{
			GST title = GST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Title_12.json");
			GST page = GST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_110_2013_Page.json");
			FilteredElementCollector collection = new FilteredElementCollector(Rvt.Handler.Doc);
			ElementSet set = GostTools.ElementConvert(collection.OfCategory(BuiltInCategory.OST_PlumbingFixtures).ToElements());
			page.AddElement(set);

			GST stamp = GST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Stamp_3.json");
			GST dop = GST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Dop_3.json");

			WorkBook wb = new WorkBook();
			WorkSheet wst = new WorkSheet(wb, "Титул");
			WorkSheet ws = new WorkSheet(wb);
			wst.AddTable(title);
			ws.AddTable(page);
			ws.AddTable(stamp);
			ws.AddTable(dop);
			wst.BuildWorkSheet();
			ws.BuildWorkSheet();
			wb.CloseWorkBook();
		}

		//internal void CreateFileTest(Document doc)
		//{
		//	WorkBook workbook = new WorkBook();
		//	GOST_21_110_2013 gost = new GOST_21_110_2013(workbook, "Мебель");

		//	FilteredElementCollector collection = new FilteredElementCollector(doc);
		//	List<Element> list = collection.OfCategory(BuiltInCategory.OST_Furniture).ToElements().Take(8).ToList<Element>();
		//	//List<Element> list_e = list.ToList<Element>();

		//	//gost.FillLines(list);
		//	//gost.FillTable();

		//	workbook.CloseWorkBook();
		//}

		//internal void CreateLogTest(Document document)
		//{
		//	WorkBook workbook = new WorkBook();
		//	ParamsXml prm = new ParamsXml(workbook);

		//	//FilteredElementCollector collection = new FilteredElementCollector(doc);
		//	prm.GetElements(document);

		//	workbook.CloseWorkBook();
		//}

		//internal void TableShapeTest(Document document)
		//{
		//	WorkBook workbook = new WorkBook();
		//	//ParamsXml prm = new ParamsXml(workbook);

		//	GostTools a = new AbstractGOST(workbook);

		//	workbook.CloseWorkBook();
		//}


		/*
		IList<Element> getCollectionByCategory(Document doc, BuiltInCategory category)
		{
			FilteredElementCollector collection = new FilteredElementCollector(doc);

			IList<Element> list = collection.OfCategory(category).ToElements();
			return list;
		}
		*/

	} // public class MainClass : IExternalCommand

	//public class Utils
	//{

	//} // public class Utils

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
		};
		public struct A4
		{
			public const int Height = 297;
			public const int Width = 210;
		};

		// to delete::::
		public struct A11
		{
			public const int Height = 100;
			public const int Width = 60;
		};
	} // static class Constants

	public class ElementsContainer
	{
		public List<ElementCollection> Collections;
		public Array CategoriesList = Enum.GetValues(typeof(BuiltInCategory));
		public Array ParametersList = Enum.GetValues(typeof(BuiltInParameter));
	}

	public class ElementCollection
	{
		//private List<List<string>> _listOfCategories;
		public string CategoryName { get; set; }
		public List<Element> Elements;
	}

} // namespace RevToGOSTv0
