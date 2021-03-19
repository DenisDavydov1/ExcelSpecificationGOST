using System;
using System.Collections.Generic;
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
	abstract class GOST
	{

	/*
	**	Member properties
	*/

		public WorkBook Workbook { get; set; }
		public int Format { get; set; }
		public int Orientation { get; set; }
		public int Stamp { get; set; }
		//public List<IXLWorksheet> Worksheets;
		public string WorksheetName { get; set; }
		public IXLWorksheet Worksheet { get; set; }

	/*
	**	Member fields
	*/
	
		protected List< List<string> > Lines;

	/*
	**	Member enums
	*/

		public enum Formats { A3, A4, A5 }
		public enum Orientations { Portrait, Landscape }

	/*
	**	Member methods
	*/

		public abstract void SetFormat();
		public abstract void FormatHeader();
		//public abstract void CreateWorksheet(IXLWorkbook Workbook); // made in constructor

		// public abstract bool IsEmptyLine(int line);
		// public abstract int GetFreeLine(int line);
		// public abstract bool IsAvailableList(int line);

		public abstract void FillLines(List<Element> elements);
		public abstract void FillTable();


	} // class GOST
} // namespace RevToGOSTv0
