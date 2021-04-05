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
		** Member fields
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

			

			Work.Gost.AddElement(Rvt.Data.PickedElements);

			//Work.Book = new WorkBook();
			Work.Book.AddWorkSheet("Листик");
			Work.Book.WSs[0].AddTable(Work.Gost);

			GOST stamp = GOST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Stamp_3.json");
			GOST dop = GOST.LoadConfFile(@"F:\CS_CODE\REVIT\PROJECTS\Templates\GOST_21_101_2020_Dop_3.json");
			Work.Book.WSs[0].AddTable(stamp);
			Work.Book.WSs[0].AddTable(dop);

			Work.Book.BuildWorkSheets();
			Work.Book.SetWorkbookAuthor();

			SaveProcedure();
		}

		private void SaveProcedure()
		{
			SaveFileDialog saveFileDialog = new SaveFileDialog();
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
						return;
					}
				}
				else
					return;
			}
		}

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
