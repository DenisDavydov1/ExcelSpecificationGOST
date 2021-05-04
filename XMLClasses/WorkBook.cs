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
		#region properties

		public XLWorkbook WB { get; set; }
		public List<WorkSheet> WSs { get; set; }
		
		public GOST.Standarts Title { get; set; }
		public GOST.Standarts Table { get; set; }
		public GOST.Standarts Stamp1 { get; set; }
		public GOST.Standarts Stamp2 { get; set; }
		public GOST.Standarts Dop1 { get; set; }
		public GOST.Standarts Dop2 { get; set; }

		public int Pages { get; set; } = 0;

		#endregion properties

		#region methods

		public WorkBook()
		{
			InitWorkBook();
			Title = GOST.Standarts.None;
			Table = GOST.Standarts.None;
			Stamp1 = GOST.Standarts.None;
			Dop1 = GOST.Standarts.None;
			Stamp2 = GOST.Standarts.None;
			Dop2 = GOST.Standarts.None;
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
			WB.Properties.Author = "RevitToGOST";
			WB.Properties.Title = Rvt.Handler.ProjInfo.Name;
			// WB.Properties.Subject = "theSubject";
			// WB.Properties.Category = "theCategory";
			// WB.Properties.Keywords = "theKeywords";
			// WB.Properties.Comments = "theComments";
			WB.Properties.Status = Rvt.Handler.ProjInfo.Status;
			WB.Properties.LastModifiedBy = Rvt.Handler.ProjInfo.Author;
			WB.Properties.Company = Rvt.Handler.ProjInfo.OrganizationName;
			// WB.Properties.Manager = "Денис Давыдов";
		}

		public void BuildWorkSheets()
		{
			foreach (WorkSheet ws in Work.Book.WSs)
			{
				if (Rvt.Control.ExportWorker.CancellationPending == false)
				{
					ws.BuildWorkSheet();
					Rvt.Control.Progress += (95 - 15) / Work.Book.WSs.Count;
				}
				else
					return;
			}
		}

		public void InitWorkBook()
		{
			WB = new XLWorkbook();
			WSs = new List<WorkSheet>();
			Pages = 0;
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

					if (Pages == 0)
					{
						// Add stamp to page
						if (Stamp1 != GOST.Standarts.None)
							newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Stamp1]));

						// Add dop to page
						if (Dop1 != GOST.Standarts.None)
							newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Dop1]));
					}
					else
					{
						// Add stamp to page
						if (Stamp2 != GOST.Standarts.None)
							newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Stamp2]));

						// Add dop to page
						if (Dop2 != GOST.Standarts.None)
							newWS.AddTable(GOST.LoadConfFile(ConfFile.Conf[(int)Dop2]));
					}

					Pages++;
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
			if (ConfFile.FillTitle[(int)Title] == null)
				return;
			ConfFile.FillTitle[(int)Title](0);
		}

		public void FillStamps()
		{
			for (int i = 0; i < Work.Book.Pages; ++i)
			{
				if (i == 0)
				{
					if (Work.Book.Stamp1 == GOST.Standarts.None || ConfFile.FillStamp[(int)Stamp1] == null)
						continue;
					ConfFile.FillStamp[(int)Stamp1](i);
				}
				else
				{
					if (Work.Book.Stamp2 == GOST.Standarts.None || ConfFile.FillStamp[(int)Stamp2] == null)
						continue;
					ConfFile.FillStamp[(int)Stamp2](i);
				}
			}
		}

		public void FillDops()
		{
			for (int i = 0; i < Work.Book.Pages; ++i)
			{
				if (i == 0)
				{
					if (Work.Book.Dop1 == GOST.Standarts.None || ConfFile.FillDop[(int)Dop1] == null)
						continue;
					ConfFile.FillDop[(int)Dop1](i);
				}
				else
				{
					if (Work.Book.Dop2 == GOST.Standarts.None || ConfFile.FillDop[(int)Dop2] == null)
						continue;
					ConfFile.FillDop[(int)Dop2](i);
				}
			}
		}

		#endregion methods

	} // class WorkBook

} // namespace RevitToGOST
