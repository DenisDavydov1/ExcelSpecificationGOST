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

	/*
	**	Member properties
	*/

		public string FilePath { get; set; }
		public IXLWorkbook Workbook { get; set; }

	/*
	**	Member methods
	*/

		public WorkBook(string filePath = Constants.DefaultFilePath)
		{
			this.FilePath = filePath;
			this.CreateWorkBook();
		}

		public void CreateWorkBook()
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
			this.Workbook.Properties.Author = "RevToGOSTv0";
			this.Workbook.Properties.Title = "Спецификация";
			this.Workbook.Properties.Subject = "theSubject";
			this.Workbook.Properties.Category = "theCategory";
			this.Workbook.Properties.Keywords = "theKeywords";
			this.Workbook.Properties.Comments = "theComments";
			this.Workbook.Properties.Status = "theStatus";
			this.Workbook.Properties.LastModifiedBy = "RevToGOSTv0";
			this.Workbook.Properties.Company = "Денис Давыдов";
			this.Workbook.Properties.Manager = "Денис Давыдов";
		}

	} // class Workbook

} // namespace RevToGOSTv0
