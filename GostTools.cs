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

	static class ElementExtension
	{
		public static string Fuck(this Element elem)
		{
			return "Fuck";
		}

		public static string InstanceName(this Element elem)
		{
			Log.WriteLine("Getting instance name: {0}", Work.Gost.ElementInstanceName(elem));
			return Work.Gost.ElementInstanceName(elem);
		}
	}

} // namespace RevitToGOST
