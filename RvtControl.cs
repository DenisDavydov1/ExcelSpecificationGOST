using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace RevitToGOST
{
	static partial class Rvt
	{
		public static RvtControl Control;
	}

	public class Progress : INotifyPropertyChanged
	{
		public double Value { get; set; } = 0.0;
		public event PropertyChangedEventHandler PropertyChanged;
	}

	class RvtControl
	{
		/*
		** Member properties
		*/

		public bool GroupElemsCheckBox { get; set; } = false;
		public bool EnumerateColumnsCheckBox { get; set; } = false;
		public double ExportProgressBarValue { get; set; } = 0.0;
		public Progress Progress { get; set; } = new Progress();

		/*
		** Member methods
		*/

		public void ExportButton()
		{
			//Rvt.Progress = new ExportProgress();
			//ExportProgress progress = new ExportProgress();
			//MainWindow.SetProgressValue(0.0);

			// Pick data and assign it to Rvt.Data.ExportElements
			Rvt.Data.SetExportElements();

			//Rvt.Progress.Value = 10.0;

			// Enumerate columns - add lines to Rvt.Data.ExportElements (if a box checked)
			Rvt.Data.InsertColumnsEnumerationLines();

			//Rvt.Progress.Value = 20.0;

			// Load needed configuration files and add worksheets
			// and assign worksheets with GOSTs
			Work.Book.LoadConfigs();

			//Rvt.Progress.Value = 30.0;

			// Fill tables with element collection
			Rvt.Data.FillLines();
			Work.Book.AddExportElements();
			Work.Book.ConvertElementCollectionsToLists();

			//Rvt.Progress.Value = 40.0;

			// Fill stamps
			// Fill dops
			// Fill title page
			Work.Book.FillTitlePage();

			//Rvt.Progress.Value = 60.0;

			// Build tables and fill it with data lines
			Work.Book.BuildWorkSheets();

			//Rvt.Progress.Value = 80.0;

			// Move title page to front of workbook
			Work.Book.MovePagesToRightPlaces();

			//Rvt.Progress.Value = 90.0;

			// Set author parameters to workbook
			Work.Book.SetWorkbookAuthor();

			//Rvt.Progress.Value = 100.0;

			// Save file
			SaveProcedure();

			// Unset Rvt.Data.ExportElements
			Rvt.Data.InitExportElements();
		}

		private void SaveProcedure()
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
