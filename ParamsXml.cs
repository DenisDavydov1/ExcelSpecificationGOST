//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using ClosedXML.Excel;

//using Autodesk.Revit.DB;
//using Autodesk.Revit.DB.Architecture;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Selection;
//using Autodesk.Revit.ApplicationServices;
//using Autodesk.Revit.Attributes;

//namespace RevToGOSTv0
//{
//	class ParamsXml : GostData
//	{

//	/*
//	**	Member properties
//	*/

//		private List<string> Header;

//	/*
//	**	Member methods
//	*/

//		public ParamsXml(WorkBook workbook, string worksheetName = Constants.DefaultName)
//		{
//			this.WorksheetName = worksheetName;
//			this.Workbook = workbook;
//			this.Worksheet = this.Workbook.AddWorkSheet(this.WorksheetName);

//			this.SetFormat();
//			this.FormatHeader();

//			this.Lines = new List<List<string>>();
//			this.Header = new List<string>();
//		}

//		public void GetElements(Document document)
//		{
//			foreach (BuiltInCategory category in Enum.GetValues(typeof(BuiltInCategory)))
//			{
//				FilteredElementCollector collection = new FilteredElementCollector(document);
//				try
//				{
//					string key = Enum.GetName(typeof(BuiltInCategory), category);
//					Log.WriteLine(String.Format("Category name: {0} {1}\n", key, (int)category));
//					List<Element> list = collection.OfCategory(category).ToElements().ToList<Element>();
//					Log.WriteLine(String.Format("Elements count: {0}", list.Count));
//					this.FillLines(list);
//				}
//				catch (Exception e)
//				{
//					Log.WriteLine(String.Format("EXCEPTION: {0}", e.Message));
//				}
//			}
//		}

//		public override void FillLines(List<Element> elements)
//		{
//			foreach (Element elem in elements)
//			{
//				ParameterMap map = elem.ParametersMap;
//				ParameterMapIterator it = map.ForwardIterator();
//				while (it.MoveNext())
//				{
//					Parameter p = (Parameter)it.Current;
//					Log.WriteLine(string.Format("{0} : AsDouble={1}, AsElementId={2}, AsInteger={3}, AsString={4}, AsValueString={5}", it.Key ?? "null",
//						p.AsDouble(), p.AsElementId(), p.AsInteger(), p.AsString(), p.AsValueString()));
//				}
//				Log.WriteLine("____________________");
//			}
//			Log.Write("\n\n\n\n");
//		}

//		public override void FillTable()
//		{
//		}

//		public override void FormatHeader()
//		{
//		}

//		public override void SetFormat()
//		{
//			//// Page setup
//			//this.Worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
//			//this.Worksheet.PageSetup.AdjustTo(100);
//			//this.Worksheet.PageSetup.PaperSize = XLPaperSize.A3Paper;
//			//this.Worksheet.PageSetup.VerticalDpi = 600;
//			//this.Worksheet.PageSetup.HorizontalDpi = 600;

//			//// Margins setup
//			//this.Worksheet.PageSetup.Margins.Top = 0.5 / Constants.inch;
//			//this.Worksheet.PageSetup.Margins.Left = 2 / Constants.inch;
//			//this.Worksheet.PageSetup.Margins.Right = 0.5 / Constants.inch;
//			//this.Worksheet.PageSetup.Margins.Bottom = 0.5 / Constants.inch;
//			//this.Worksheet.PageSetup.Margins.Footer = 0;
//			//this.Worksheet.PageSetup.Margins.Header = 0;
//			//this.Worksheet.PageSetup.CenterHorizontally = false;
//			//this.Worksheet.PageSetup.CenterVertically = false;
//		}
//	}
//}
