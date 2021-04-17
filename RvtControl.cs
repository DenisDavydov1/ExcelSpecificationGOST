using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RevitToGOST
{
	static partial class Rvt
	{
		public static RvtControl Control;
	}

	class RvtControl
	{
		#region properties

		///// Control elements values /////
		public bool GroupElemsCheckBox { get; set; } = false;
		public bool EnumerateColumnsCheckBox { get; set; } = false;

		///// Application exception container /////
		public Exception LastException { get; set; } = new Exception();

		//// Progress bar value ////
		private int _Progress;  // for Progress Bar
		public int Progress
		{
			get { return _Progress; }
			set
			{
				if (value != _Progress)
				{
					_Progress = value;
					if (ExportWorker != null)
						ExportWorker.ReportProgress(Progress);
				}
			}
		}

		//// Background export thread ////
		public BackgroundWorker ExportWorker { get; set; } = null;

		#endregion properties

		#region methods

		private bool CheckForCancellation(DoWorkEventArgs e)
		{
			if (ExportWorker.CancellationPending == true)
			{
				e.Cancel = true;
				return true;
			}
			return false;
		}

		public void ExportProcedure(object sender, DoWorkEventArgs e)
		{
			ExportWorker = sender as BackgroundWorker;
			Progress = 0; if (CheckForCancellation(e) == true) return;

			Work.Book.InitWorkBook();
			// Pick data and assign it to Rvt.Data.ExportElements
			Rvt.Data.SetExportElements();

			// Enumerate columns - add lines to Rvt.Data.ExportElements (if a box checked)
			Rvt.Data.InsertColumnsEnumerationLines();

			// Load needed configuration files and add worksheets
			// and assign worksheets with GOSTs
			Work.Book.LoadConfigs();

			Progress = 5; if (CheckForCancellation(e) == true) return;

			// Fill tables with element collection
			Rvt.Data.FillLines();
			Work.Book.AddExportElements();
			Work.Book.ConvertElementCollectionsToLists();

			Progress = 10; if (CheckForCancellation(e) == true) return;

			// Fill stamps
			Work.Book.FillStamps();
			// Fill dops
			Work.Book.FillDops();
			// Fill title page
			Work.Book.FillTitlePage();

			Progress = 15; if (CheckForCancellation(e) == true) return;

			// Build tables and fill it with data lines
			Work.Book.BuildWorkSheets();

			Progress = 95; if (CheckForCancellation(e) == true) return;

			// Move title page to front of workbook
			Work.Book.MovePagesToRightPlaces();

			// Set author parameters to workbook
			Work.Book.SetWorkbookAuthor();

			Progress = 100; if (CheckForCancellation(e) == true) return;
			ExportWorker = null;
		}

		public void SaveProcedure()
		{
			// New save dialog window
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.Filter = "Книга Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";

			// Set .rvt directory as initial for save procedure
			string savePath = Rvt.Handler.Doc.PathName;
			savePath = savePath.Substring(0, savePath.LastIndexOf(@"\"));
			saveFileDialog.InitialDirectory = savePath;

			// Save loop
			while (true)
			{
				if (saveFileDialog.ShowDialog() == true)
				{
					if (IsFileLocked(saveFileDialog.FileName) == true)
					{
						MessageBox.Show("Файл занят другим процессом. Закройте файл в другой программе и попробуйте снова.",
							"Ошибка сохранения", MessageBoxButton.OK, MessageBoxImage.Error);
					}
					else
					{
						Work.Book.SaveAs(saveFileDialog.FileName);
						MessageBox.Show("Таблица успешно экспортирована.\nПуть к файлу: " + saveFileDialog.FileName,
							"Готово", MessageBoxButton.OK);
						Work.Book.InitWorkBook();
						return;
					}
				}
				else
				{
					Work.Book.InitWorkBook();
					return;
				}
			}
		}

		// Checks if a file is ready for writing
		private bool IsFileLocked(string filePath)
		{
			try
			{
				using (FileStream stream = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.Write))
					stream.Close();
			}
			catch (IOException)
			{
				return true;
			}
			return false;
		}

		#endregion methods

	} // class RvtControl

} // namespace RevitToGOST
