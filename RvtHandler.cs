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

namespace RevToGOSTv0
{
	static class Rvt
	{
		public static RvtHandler Handler;
	}

	class RvtHandler
	{
		/*
		** Member fields
		*/

		public ExternalCommandData ExtData { get; set; }
		public UIApplication UIApp { get; set; }
		public Document Doc { get; set; }
		public ProjectInfo ProjInfo { get; set; }
		//public List<int>


		/*
		** Member methods
		*/

		public RvtHandler(ExternalCommandData extData, ElementSet elemSet)
		{
			ExtData = extData;
			UIApp = ExtData.Application;
			Doc = UIApp.ActiveUIDocument.Document;
			ProjInfo = Doc.ProjectInformation;
		}

		public void LogCategories()
		{
			foreach (BuiltInCategory category in Enum.GetValues(typeof(BuiltInCategory)))
			{
				FilteredElementCollector collection = new FilteredElementCollector(Doc);
				try
				{
					string key = Enum.GetName(typeof(BuiltInCategory), category);
					//List<Element> list = collection.OfCategory(category).ToElements().ToList<Element>();
					IList<Element> list = collection.OfCategory(category).ToElements();
					if (list.Count > 0)
					{
						Log.WriteLine(String.Format("Category name: {0, 60} [{1}]. Elements count: {2, 3}", key, (int)category, list.Count));
					}
				}
				catch (Exception e)
				{
					Log.WriteLine(String.Format("EXCEPTION: {0}", e.Message));
				}
			}
		}

		public void LogCategory(BuiltInCategory category)
		{
			FilteredElementCollector collection = new FilteredElementCollector(Doc);
			IList<Element> list = collection.OfCategory(category).ToElements();
			//for (int i = 0; i < list.Count; ++i)
			for (int i = 0; i < 10; ++i)
			{
				foreach (BuiltInParameter param in Enum.GetValues(typeof(BuiltInParameter)))
				{
					//string val = list[i].get_Parameter(BuiltInParameter.)
					string key = Enum.GetName(typeof(BuiltInParameter), param);
					Parameter prm = list[i].get_Parameter(param);
					if (key == null || prm == null)
						continue;
					string asStr = prm.AsString();
					string avs = prm.AsValueString();
					double asDbl = prm.AsDouble();
					if ((asStr == null || asStr == "") && (avs == null || avs == "") && asDbl == 0.0)
						continue;
					Log.WriteLine("Param name: {0, 30} [{1}]\nValue as string: {2}\nValue as value string: {3}\nValue as double: {4}",
						key, (int)param, prm.AsString(), prm.AsValueString(), prm.AsDouble());
					Log.WriteLine("\n");
				}
				//IList<Parameter> param = list[i].GetOrderedParameters();
				//for (int j = 0; j < param.Count; ++j)
				//{
				//	Log.WriteLine("Parameter: {0, 40}. Value: {1, 20}", param[j].Definition.Name, param[j].AsValueString());
				//}
				Log.WriteLine("\n\n\n");
			}
		}

		private static bool InParamSet(Parameter param, ParameterSet set)
		{
			ParameterSetIterator it = set.ForwardIterator();
			while (it.MoveNext())
			{
				try
				{
					if (param.Definition.Name == ((Parameter)it.Current).Definition.Name)
						return true;
				}
				catch (Exception e) { continue; }
			}
			return false;
		}

		public void LogAllElements()
		{
			foreach (BuiltInCategory category in Enum.GetValues(typeof(BuiltInCategory)))
			{
				FilteredElementCollector collection = new FilteredElementCollector(Doc);
				IList<Element> list = new List<Element>();
				try {
					list = collection.OfCategory(category).ToElements();
				} catch (Exception e) { continue; }
				if (list.Count == 0)
					continue;
				Log.WriteLine("\n\nNEW CATEGORY: {0}\n\n", Enum.GetName(typeof(BuiltInCategory), category));
				int i = 0;
				foreach (Element elem in list)
				{
					ParameterSet prms = elem.Parameters;
					foreach (BuiltInParameter param in Enum.GetValues(typeof(BuiltInParameter)))
					{
						Parameter prm = elem.get_Parameter(param);
						if (InParamSet(prm, prms) == false)
							continue;
						try	{
							Log.WriteLine("[Name] {4, 40} [Val] {5, 7} [Descr] {0, 40} [String] {1, 40} [ValueStr] {2, 40} [Double] {3, 10}",
							prm.Definition.Name, prm.AsString(), prm.AsValueString(), prm.AsDouble(), Enum.GetName(typeof(BuiltInParameter), param), (int)param);
						} catch (Exception e) { continue; }
						//Log.WriteLine("\n\nNEW CATEGORY: {0}\n\n", Enum.GetName(typeof(BuiltInCategory), category));
					}
					//ParameterSetIterator it = prms.ForwardIterator();
					//while (it.MoveNext())
					//{
					//	try
					//	{
					//		Log.WriteLine("[Name] {0, 40} [String] {1, 40} [ValueStr] {2, 40} [Double] {3, 10}",
					//		((Parameter)it.Current).Definition.Name, ((Parameter)it.Current).AsString(), ((Parameter)it.Current).AsValueString(), ((Parameter)it.Current).AsDouble());
					//	} catch (Exception e) { continue; }
					//}
					Log.WriteLine("");
					i++;
					if (i == 5)
						break;
				}
			}
			//foreach (BuiltInCategory category in Enum.GetValues(typeof(BuiltInCategory)))
			//{
			//	FilteredElementCollector collection = new FilteredElementCollector(Doc);
			//	IList<Element> list = collection.OfCategory(category).ToElements();
			//	foreach (Element elem in list)
			//	{
			//		foreach
			//	}
			//}
		}
	} // class RvtHandler
} // namespace RevToGOSTv0
