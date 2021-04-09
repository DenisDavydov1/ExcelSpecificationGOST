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
	static partial class Work
	{
		public static WorkBook Book { get; set; }
	}

	public class WorkBook
	{
		/*
		**	Member fields
		*/

		//public string FilePath { get; set; }
		public XLWorkbook WB { get; set; }
		public List<WorkSheet> WSs { get; set; }
		
		public GOST.Standarts Title { get; set; }
		public GOST.Standarts Table { get; set; }
		public GOST.Standarts Stamp { get; set; }
		public GOST.Standarts Dop { get; set; }


		/*
		**	Member methods
		*/

		public WorkBook()
		{
			InitWorkBook();
			Title = GOST.Standarts.None;
			Table = GOST.Standarts.None;
			Stamp = GOST.Standarts.None;
			Dop = GOST.Standarts.None;
		}

		public void SaveAs(string filePath)
		{
			WB.SaveAs(filePath);
		}
		
		public WorkSheet AddWorkSheet(string worksheetName, int position = -1)
		{
			if (position == -1)
			{
				IXLWorksheet ixlWorkSheet = WB.Worksheets.Add(worksheetName);
				WorkSheet newWS = new WorkSheet(ixlWorkSheet, worksheetName);
				WSs.Add(newWS);
				return newWS;
			}
			else
			{
				IXLWorksheet ixlWorkSheet = WB.Worksheets.Add(worksheetName, position);
				WorkSheet newWS = new WorkSheet(ixlWorkSheet, worksheetName);
				WSs.Insert(position, newWS);
				return newWS;
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
			{
				ws.BuildWorkSheet();
			}
		}

		public void InitWorkBook()
		{
			WB = new XLWorkbook();
			WSs = new List<WorkSheet>();
		}

		private static int GetTablePagesCount(int elems, int lines)
		{
			if (elems < 1 || lines < 1)
				return 1;
			return elems / lines + (elems % lines == 0 ? 0 : 1);
		}

		public void LoadConfigs()
		{
			// Add table page(s)
			if (Table != GOST.Standarts.None)
			{
				for (int i = 0; i < GetTablePagesCount(Rvt.Data.ExportElements.Count, ConfFile.Lines[(int)Table]); ++i)
				{
					WorkSheet newWS = AddWorkSheet(String.Format("Лист {0}",  i + 1));
					newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Table]));
						
					// Add stamp to page
					if (Stamp != GOST.Standarts.None)
					{
						// TO DO HERE! Проверка, если это второй лист, надо добавить другую основную надпись и доп графу
						newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Stamp]));
					}

					// Add dop to page
					if (Dop != GOST.Standarts.None)
					{
						// HERE TOO!
						newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Dop]));
					}
				}
			}

			// Add title page
			if (Title != GOST.Standarts.None)
			{
				WorkSheet newWS = AddWorkSheet("Титульный лист");
				newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Title]));
			}
		}

		public void AddExportElements()
		{
			int totalLines = Rvt.Data.ExportElements.Count;
			for (int line = 0, page = 0; line < totalLines; line++)
			{
				if (WSs[page].Tables[0].IsFull == true)
					page++;
				WSs[page].Tables[0].AddElement(Rvt.Data.ExportElements[line]);
			}
		}

		public void ConvertElementCollectionsToLists()
		{
			foreach (WorkSheet ws in WSs)
				ws.TablesElemColsToData();
		}

		public void MovePagesToRightPlaces()
		{
			if (Title != GOST.Standarts.None)
				WB.Worksheet(WB.Worksheets.Count()).Position = 1;
		}

		public void FillTitlePage()
		{
			if (Title == GOST.Standarts.None)
				return;
			if (ConfFile.FillLine[(int)Title] == null)
				return;
			ConfFile.FillLine[(int)Title](null);
		}

	} // class WorkBook

} // namespace RevitToGOST
