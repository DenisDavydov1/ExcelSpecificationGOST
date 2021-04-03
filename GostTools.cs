using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

namespace RevToGOSTv0
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
} // namespace RevToGOSTv0
