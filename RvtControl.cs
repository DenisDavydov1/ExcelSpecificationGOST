using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
	class RvtControl
	{
		/*
		** Member properties
		*/

		public bool GroupElemsCheckBox { get; set; } = false;
		public bool EnumerateColumnsCheckBox { get; set; } = false;

		/*
		** Member methods
		*/

		public void ExportButton()
		{
			// Pick data and assign it to Rvt.Data.ExportElements
			Rvt.Data.SetExportElements();

			// Enumerate columns - add lines to Rvt.Data.ExportElements (if a box checked)
			Rvt.Data.InsertColumnsEnumerationLines();

			// Load needed configuration files and add worksheets
			// and assign worksheets with GOSTs
			Work.Book.LoadConfigs();

			// Add element collection to GOSTs
			Rvt.Data.CreateLines();
			Work.Book.AddExportElements();

			// Build tables and fill it with data lines
			Work.Book.BuildWorkSheets();

			// Move title page to front of workbook
			Work.Book.MovePagesToRightPlaces();

			// Set author parameters to workbook
			Work.Book.SetWorkbookAuthor();

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
