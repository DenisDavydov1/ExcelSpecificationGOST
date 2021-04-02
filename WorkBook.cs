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
	static class Work
	{
		public static WorkBook Book { get; set; }
	}
	class WorkBook
	{
		/*
		**	Member fields
		*/

		public string FilePath { get; set; }
		public IXLWorkbook WB { get; set; }
		public List<IXLWorksheet> WSs { get; set; }


		/*
		**	Member methods
		*/

		public WorkBook(string filePath = Constants.DefaultFilePath)
		{
			FilePath = filePath;
			WB = new XLWorkbook();			
		}
				
		public void CloseWorkBook()
		{
			SetWorkbookAuthor();
			WB.SaveAs(FilePath);
		}

		public IXLWorksheet AddWorkSheet(string worksheetName, int position = -1)
		{
			IXLWorksheet worksheet;

			if (position == -1)
			{
				worksheet = WB.Worksheets.Add(worksheetName);
				WSs.Add(worksheet);
			}
			else
			{
				worksheet = WB.Worksheets.Add(worksheetName, position);	// mb error in position indexing in workbook and list
				WSs.Insert(position, worksheet);
			}
			return worksheet;
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

	} // class Workbook

} // namespace RevToGOSTv0
