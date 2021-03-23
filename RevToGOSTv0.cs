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

			UIApplication uiApp = commandData.Application;
			Document doc = uiApp.ActiveUIDocument.Document;

			//CreateFileTest(doc);
			//CreateLogTest(doc);
			TableShapeTest(doc);

			//List<string> list_s = new List<string>();
			//var enumvar = Enum.GetValues(typeof(BuiltInCategory));
			//foreach (var name in enumvar)
			//{
			//	System.Diagnostics.Debug.WriteLine(name.ToString());
			//	list_s.Add(name.ToString());
			//}
			//List< List<string> > list_s = new List<List<string>>();

			//list_s.Add(GOST1.GetTableLine(elem));

			//foreach (var v in list_s)
			//Utils.PrintList(Constants.logpath, v);

			//Utils.PrintList(Constants.logpath, list_s);

			//uiWindow uiWin = new uiWindow(list);
			//uiWin.Show();

			return Result.Succeeded;
		}

		internal void CreateFileTest(Document doc)
		{
			WorkBook workbook = new WorkBook();
			GOST_21_110_2013 gost = new GOST_21_110_2013(workbook, "Мебель");

			FilteredElementCollector collection = new FilteredElementCollector(doc);
			List<Element> list = collection.OfCategory(BuiltInCategory.OST_Furniture).ToElements().Take(8).ToList<Element>();
			//List<Element> list_e = list.ToList<Element>();

			//gost.FillLines(list);
			//gost.FillTable();

			workbook.CloseWorkBook();
		}

		internal void CreateLogTest(Document document)
		{
			WorkBook workbook = new WorkBook();
			ParamsXml prm = new ParamsXml(workbook);

			//FilteredElementCollector collection = new FilteredElementCollector(doc);
			prm.GetElements(document);

			workbook.CloseWorkBook();
		}

		internal void TableShapeTest(Document document)
		{
			WorkBook workbook = new WorkBook();
			//ParamsXml prm = new ParamsXml(workbook);

			AbstractGOST a = new AbstractGOST(workbook);

			workbook.CloseWorkBook();
		}


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
		public const string LogPath = "F:\\CS_CODE\\REVIT\\PROJECTS\\tmp";
		public const string DefaultFilePath = "F:\\CS_CODE\\REVIT\\PROJECTS\\ClosedXMLTests\\output\\wb.xlsx";
		public const double inch = 2.54;
		public const double mm_w = 0.483;
		public const double mm_h = 2.9;
		public const string DefaultName = "Без названия";
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
