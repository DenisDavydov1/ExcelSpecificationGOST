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

		public bool GroupElemsCheckBox { get; set; }

		/*
		** Member methods
		*/

		public RvtControl()
		{
			GroupElemsCheckBox = false;
		}

		public void ExportButton()
		{
			//GOST page = GOST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_110_2013_Table1.json");
			//GOST stamp = GOST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Stamp3.json");
			//GOST dop = GOST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Dop3.json");

			//page.AddElement(Rvt.Data.PickedElements);

			Work.Book.LoadConfigs();
			Work.Book.AddElementCollection(Rvt.Data.PickedElements);

			Work.Book.BuildWorkSheets();
			Work.Book.SetWorkbookAuthor();

			SaveProcedure();
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

		private bool IsFileLocked(string filePath)
		{
			// Check if a file is ready for writing
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
