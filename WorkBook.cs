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

namespace RevitToGOST
{
	static class Work
	{
		public static WorkBook Book { get; set; }
		public static GOST Gost { get; set; }
	}

	class WorkBook
	{
		/*
		**	Member fields
		*/

		public string FilePath { get; set; }
		public XLWorkbook WB { get; set; }
		public List<WorkSheet> WSs { get; set; }


		/*
		**	Member methods
		*/

		public WorkBook(string filePath = Constants.DefaultFilePath)
		{
			FilePath = filePath;
			WB = new XLWorkbook();
			WSs = new List<WorkSheet>();
		}

		public void SaveAs(string filePath = Constants.DefaultFilePath)
		{
			WB.SaveAs(filePath);
		}

		public void AddWorkSheet(string worksheetName, int position = -1)
		{
			if (position == -1)
			{
				IXLWorksheet ixlWorkSheet = WB.Worksheets.Add(worksheetName);
				WorkSheet newWS = new WorkSheet(ixlWorkSheet, worksheetName);
				WSs.Add(newWS);
			}
			else
			{
				IXLWorksheet ixlWorkSheet = WB.Worksheets.Add(worksheetName, position);
				WorkSheet newWS = new WorkSheet(ixlWorkSheet, worksheetName);
				WSs.Insert(position, newWS);
			}
		}

		public void SetWorkbookAuthor()
		{
			WB.Properties.Author = "RevToGOSTv0";
			WB.Properties.Title = "Спецификация";
			WB.Properties.Subject = "theSubject";
			WB.Properties.Category = "theCategory";
			WB.Properties.Keywords = "theKeywords";
			WB.Properties.Comments = "theComments";
			WB.Properties.Status = "theStatus";
			WB.Properties.LastModifiedBy = "RevToGOSTv0";
			WB.Properties.Company = "Денис Давыдов";
			WB.Properties.Manager = "Денис Давыдов";
		}

		public void BuildWorkSheets()
		{
			foreach (WorkSheet ws in Work.Book.WSs)
				ws.BuildWorkSheet();
		}
	} // class WorkBook
} // namespace RevitToGOST
