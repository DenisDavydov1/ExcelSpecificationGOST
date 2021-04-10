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
		/*
		** Member properties
		*/

		///// Control elements values /////
		public bool GroupElemsCheckBox { get; set; } = false;
		public bool EnumerateColumnsCheckBox { get; set; } = false;

		///// Application exception container /////
		public Exception LastException { get; set; } = new Exception();

		/*
		** Member methods
		*/

		public void ExportProcedure(object sender)
		{
			(sender as BackgroundWorker).ReportProgress(0);

			//Log.WriteLine("1");
			// Pick data and assign it to Rvt.Data.ExportElements
			Rvt.Data.SetExportElements();

			(sender as BackgroundWorker).ReportProgress(10);

			//Log.WriteLine("1");
			// Enumerate columns - add lines to Rvt.Data.ExportElements (if a box checked)
			Rvt.Data.InsertColumnsEnumerationLines();

			(sender as BackgroundWorker).ReportProgress(20);

			//Log.WriteLine("3");
			// Load needed configuration files and add worksheets
			// and assign worksheets with GOSTs
			Work.Book.LoadConfigs();

			(sender as BackgroundWorker).ReportProgress(30);

			//Log.WriteLine("4");
			// Fill tables with element collection
			Rvt.Data.FillLines();
			//Log.WriteLine("5");
			Work.Book.AddExportElements();
			//Log.WriteLine("6");
			Work.Book.ConvertElementCollectionsToLists();

			(sender as BackgroundWorker).ReportProgress(40);

			//Log.WriteLine("7");
			// Fill stamps
			Work.Book.FillStamps();
			// Fill dops
			Work.Book.FillDops();
			// Fill title page
			Work.Book.FillTitlePage();

			(sender as BackgroundWorker).ReportProgress(60);

			//Log.WriteLine("8");
			// Build tables and fill it with data lines
			Work.Book.BuildWorkSheets();

			(sender as BackgroundWorker).ReportProgress(80);

			//Log.WriteLine("9");
			// Move title page to front of workbook
			Work.Book.MovePagesToRightPlaces();

			(sender as BackgroundWorker).ReportProgress(90);

			// Set author parameters to workbook
			Work.Book.SetWorkbookAuthor();

			(sender as BackgroundWorker).ReportProgress(100);
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

	} // class RvtControl

} // namespace RevitToGOST
