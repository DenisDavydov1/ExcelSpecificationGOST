using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace RevToGOSTv0
{
	class WorkBook
	{
		/*
		**	Member fields
		*/

		public string FilePath { get; set; }
		public IXLWorkbook Workbook { get; set; }

		/*
		**	Member methods
		*/

		public WorkBook(string filePath = Constants.DefaultFilePath)
		{
			FilePath = filePath;
			CreateWorkBook();
		}

		private void CreateWorkBook()
		{
			this.Workbook = new XLWorkbook();
		}
		
		public void CloseWorkBook()
		{
			this.SetWorkbookAuthor();
			this.Workbook.SaveAs(this.FilePath);
		}

		public IXLWorksheet AddWorkSheet(string worksheetName)
		{
			IXLWorksheet worksheet = this.Workbook.Worksheets.Add(worksheetName);
			return worksheet;
		}

		public void SetWorkbookAuthor()
		{
			Workbook.Properties.Author = "RevToGOSTv0";
			Workbook.Properties.Title = "Спецификация";
			Workbook.Properties.Subject = "theSubject";
			Workbook.Properties.Category = "theCategory";
			Workbook.Properties.Keywords = "theKeywords";
			Workbook.Properties.Comments = "theComments";
			Workbook.Properties.Status = "theStatus";
			Workbook.Properties.LastModifiedBy = "RevToGOSTv0";
			Workbook.Properties.Company = "Денис Давыдов";
			Workbook.Properties.Manager = "Денис Давыдов";
		}

	} // class Workbook

} // namespace RevToGOSTv0
