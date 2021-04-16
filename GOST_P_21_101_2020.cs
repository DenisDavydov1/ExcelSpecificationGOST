using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

namespace RevitToGOST
{
	class GOST_P_21_101_2020_Stamp_3
	{
		public static void FillStamp(int pageNumber)
		{
			string DocName = Rvt.Handler.ProjInfo.Name;                 // "_1" обозначение документа
			string Address = Rvt.Handler.ProjInfo.Address;              // "_2" наименование предприятия
			string BuildingName = Rvt.Handler.ProjInfo.BuildingName;	// "_3" наименование здания
			// "_4" наименование изображений, помещенных на листе
			// "_6" условное обозначение вида документации
			string PageNumber = (pageNumber + 1).ToString();					// "_7" порядковый номер листа документа
			string Pages = Work.Book.Pages.ToString();							// "_8" общее количество листов документа
			string OrganizationName = Rvt.Handler.ProjInfo.OrganizationName;    // "_9" наименование организации, разработавшей документ
			string AuthorName = Rvt.Handler.ProjInfo.Author;					// Разраб.
			// "_10" характер работы
			// "_11" фамилии лиц, подписывающих документ
			// "_12" их подписи
			string Date = DateTime.UtcNow.ToString("MM.yy");			// "_13" датa подписания документа
			// "_14" - "_19" сведения об изменениях

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "Изм.", "Кол.уч", "Лист", "№ док.", "Подп.", "Дата" },

				new List<string>() { "Разраб", AuthorName, "", Date },
				new List<string>() { "", "", "", "" },
				new List<string>() { "", "", "", "" },
				new List<string>() { "", "", "", "" },
				new List<string>() { "Н.контр.", "", "", "" },
				new List<string>() { "", "", "", "" },

				new List<string>() { DocName },
				new List<string>() { Address },
				new List<string>() { BuildingName },
				new List<string>() { "Стадия", "Лист", "Листов" },
				new List<string>() { "", PageNumber, Pages },
				new List<string>() { "", OrganizationName }
			};
		}

	} // class GOST_P_21_101_2020_Stamp_3

	class GOST_P_21_101_2020_Dop_3
	{
		public static void FillDop(int pageNumber)
		{
			string _10 = "";	// "_10" характер работы
			string _11 = "";	// "_11" фамилии лиц, подписывающих документ
			string _12 = "";    // "_12" их подписи
			string _13 = "";    // "_13" датa подписания документа
			// DateTime.UtcNow.ToString("dd.MM.yyyy");
			string _20 = "";	// "_20" инвентарный номер подлинника
			string _21 = "";	// "_21" подпись лица, принявшего подлинник на хранение
			string _22 = "";	// "_22" инвентарный номер подлинника документа

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { _13, "" },
				new List<string>() { _12, "" },
				new List<string>() { _11, "" },
				new List<string>() { _10, "" },

				new List<string>() { "Взаим. инв. №", _22 },
				new List<string>() { "Подп. и дата", _21 },
				new List<string>() { "Инв. № подп.", _20 },

				new List<string>() { "Согласовано" }
			};
		}

	} // class GOST_P_21_101_2020_Dop_3

	class GOST_P_21_101_2020_Stamp_4
	{
		public static void FillStamp(int pageNumber)
		{
			string _1 = Rvt.Handler.ProjInfo.Name;				// "_1" обозначение документа
			//string _2 = Rvt.Handler.ProjInfo.Address;         // "_2" наименование предприятия
			//string _3 = Rvt.Handler.ProjInfo.BuildingName;    // "_3" наименование здания
			//string _4 = "";                                   // "_4" наименование изображений, помещенных на листе
			string _5 = Rvt.Handler.ProjInfo.BuildingName;		// "_5" наименование изделия
			string _6 = "";                                     // "_6" условное обозначение вида документации
			string _7 = (pageNumber + 1).ToString();            // "_7" порядковый номер листа документа
			string _8 = Work.Book.Pages.ToString();             // "_8" общее количество листов документа
			string _9 = Rvt.Handler.ProjInfo.OrganizationName;	// "_9" наименование организации, разработавшей документ
			string AuthorName = Rvt.Handler.ProjInfo.Author;    // Разраб.
			string _10 = "";                                    // "_10" характер работы
			string _11 = "";                                    // "_11" фамилии лиц, подписывающих документ
			string _12 = "";                                    // "_12" их подписи
			string _13 = DateTime.UtcNow.ToString("MM.yy");     // "_13" датa подписания документа
			// "_14" - "_19" сведения об изменениях
			string _23 = "";									// "_23" обозначение материала детали
			string _24 = "";									// "_24" массa изделия
			string _25 = "";									// "_25" масштаб

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "Изм.", "Кол.уч", "Лист", "№ док.", "Подп.", "Дата" },

				new List<string>() { "Разраб", AuthorName, "", _13 },
				new List<string>() { "", "", "", "" },
				new List<string>() { _10, _11, _12, "" },
				new List<string>() { "", "", "", "" },
				new List<string>() { "Н.контр.", "", "", "" },
				new List<string>() { "", "", "", "" },

				new List<string>() { _1 },
				new List<string>() { _5 },
				new List<string>() { "Стадия", "Масса", "Масштаб" },
				new List<string>() { _6, _24, _25 },
				new List<string>() { "Лист " + _7, "Листов " + _8 },
				new List<string>() { _23, _9 }
			};
		}

	} // class GOST_P_21_101_2020_Stamp_4

	class GOST_P_21_101_2020_Dop_4
	{
		public static void FillDop(int pageNumber)
		{
			string _20 = "";    // "_20" инвентарный номер подлинника
			string _21 = "";    // "_21" подпись лица, принявшего подлинник на хранение
			string _22 = "";    // "_22" инвентарный номер подлинника документа

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { "Взаим. инв. №", _22 },
				new List<string>() { "Подп. и дата", _21 },
				new List<string>() { "Инв. № подп.", _20 }
			};
		}

	} // class GOST_P_21_101_2020_Dop_4

	class GOST_P_21_101_2020_Stamp_5
	{
		public static void FillStamp(int pageNumber)
		{
			string _1 = Rvt.Handler.ProjInfo.Name;				// "_1" обозначение документа
			// string _2 = Rvt.Handler.ProjInfo.Address;		// "_2" наименование предприятия
			// string _3 = Rvt.Handler.ProjInfo.BuildingName;	// "_3" наименование здания
			// string _4 = "";									// "_4" наименование изображений, помещенных на листе
			string _5 = Rvt.Handler.ProjInfo.BuildingName;      // "_5" наименование изделия
			string _6 = "";										// "_6" условное обозначение вида документации
			string _7 = (pageNumber + 1).ToString();			// "_7" порядковый номер листа документа
			string _8 = Work.Book.Pages.ToString();				// "_8" общее количество листов документа
			string _9 = Rvt.Handler.ProjInfo.OrganizationName;	// "_9" наименование организации, разработавшей документ
			string AuthorName = Rvt.Handler.ProjInfo.Author;	// Разраб.
			string _10 = "";									// "_10" характер работы
			string _11 = "";									// "_11" фамилии лиц, подписывающих документ
			string _12 = "";									// "_12" их подписи
			string _13 = DateTime.UtcNow.ToString("MM.yy");		// "_13" датa подписания документа
			// "_14" - "_19" сведения об изменениях
			string _27 = Rvt.Handler.ProjInfo.ClientName;		// "_27" краткое наименование организации-заказчика

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { _27, _1, _5, _9 },
				new List<string>() { "Стадия", "Лист", "Листов" },
				new List<string>() { _6, _7, _8 },

				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "Изм.", "Кол.уч", "Лист", "№ док.", "Подп.", "Дата" },

				new List<string>() { "Разраб", AuthorName, "", _13 },
				new List<string>() { "", "", "", "" },
				new List<string>() { _10, _11, _12, "" },
				new List<string>() { "Н.контр.", "", "", "" },
				new List<string>() { "", "", "", "" }
			};
		}

	} // class GOST_P_21_101_2020_Stamp_5

	class GOST_P_21_101_2020_Dop_5
	{
		public static void FillDop(int pageNumber)
		{
			string _10 = "";    // "_10" характер работы
			string _11 = "";    // "_11" фамилии лиц, подписывающих документ
			string _12 = "";    // "_12" их подписи
			string _13 = "";    // "_13" датa подписания документа
								// DateTime.UtcNow.ToString("dd.MM.yyyy");
			string _20 = "";    // "_20" инвентарный номер подлинника
			string _21 = "";    // "_21" подпись лица, принявшего подлинник на хранение
			string _22 = "";    // "_22" инвентарный номер подлинника документа

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { _13, "" },
				new List<string>() { _12, "" },
				new List<string>() { _11, "" },
				new List<string>() { _10, "" },

				new List<string>() { "Взаим. инв. №", _22 },
				new List<string>() { "Подп. и дата", _21 },
				new List<string>() { "Инв. № подп.", _20 },

				new List<string>() { "Согласовано" }
			};
		}

	} // class GOST_P_21_101_2020_Dop_5

	class GOST_P_21_101_2020_Stamp_6
	{
		public static void FillStamp(int pageNumber)
		{
			string _1 = Rvt.Handler.ProjInfo.Name;              // "_1" обозначение документа
			string _7 = (pageNumber + 1).ToString();            // "_7" порядковый номер листа документа
			// "_14" - "_19" сведения об изменениях

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { _1 },
				new List<string>() { "Лист", _7 },

				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "", "", "", "", "", "" },
				new List<string>() { "Изм.", "Кол.уч", "Лист", "№ док.", "Подп.", "Дата" }
			};
		}

	} // class GOST_P_21_101_2020_Stamp_6

	class GOST_P_21_101_2020_Dop_6
	{
		public static void FillDop(int pageNumber)
		{
			string _20 = "";    // "_20" инвентарный номер подлинника
			string _21 = "";    // "_21" подпись лица, принявшего подлинник на хранение
			string _22 = "";    // "_22" инвентарный номер подлинника документа

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { "Взаим. инв. №", _22 },
				new List<string>() { "Подп. и дата", _21 },
				new List<string>() { "Инв. № подп.", _20 }
			};
		}

	} // class GOST_P_21_101_2020_Dop_6

	class GOST_P_21_101_2020_Table_7
	{
		public static void FillLine(ElementContainer elemCont)
		{
			//// Заголовок категории
			if (elemCont.LineType == ElementContainer.ContType.Category)
			{
				elemCont.Line = new List<string>() { "", "", elemCont.CategoryLine, "", "", "" };
			}

			//// Нумерация стоблцов
			else if (elemCont.LineType == ElementContainer.ContType.ColumnsEnumeration)
			{
				elemCont.Line = new List<string>() { "1", "2", "3", "4", "5", "6" };
			}

			//// Настоящий элемент
			else
			{
				elemCont.Line = new List<string>();

				// "Поз."
				elemCont.Line.Add(elemCont.Position.ToString());

				// Обозначение
				elemCont.Line.Add(GOST_21_110_2013.ElementType(elemCont));

				// Наименование
				elemCont.Line.Add(GOST_21_110_2013.ElementName(elemCont));

				// Кол.
				double amount = GOST_21_110_2013.ElementAmount(elemCont);
				if (amount == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(amount.ToString());

				// Масса ед., кг
				double weight = GOST_21_110_2013.ElementWeight(elemCont);
				if (weight == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(weight.ToString());

				// Примечание
				elemCont.Line.Add(GOST_21_110_2013.ElementNote(elemCont));
			}
		}
	}

	class GOST_P_21_101_2020_Table_8
	{
		public static void FillLine(ElementContainer elemCont)
		{
			//// Заголовок категории
			if (elemCont.LineType == ElementContainer.ContType.Category)
			{
				elemCont.Line = new List<string>() { "", "", elemCont.CategoryLine, "", "", "", "", "", "", "", "", "", "", "" };
			}

			//// Нумерация стоблцов
			else if (elemCont.LineType == ElementContainer.ContType.ColumnsEnumeration)
			{
				elemCont.Line = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14" };
			}

			//// Настоящий элемент
			else
			{
				elemCont.Line = new List<string>();

				// "Поз."
				elemCont.Line.Add(elemCont.Position.ToString());

				// Обозначение
				elemCont.Line.Add(GOST_21_110_2013.ElementType(elemCont));

				// Наименование
				elemCont.Line.Add(GOST_21_110_2013.ElementName(elemCont));

				// Кол.
				for (int i = 0; i < 8; ++i)
					elemCont.Line.Add(String.Empty);
				double amount = GOST_21_110_2013.ElementAmount(elemCont);
				if (amount == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(amount.ToString());

				// Масса ед., кг
				double weight = GOST_21_110_2013.ElementWeight(elemCont);
				if (weight == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(weight.ToString());

				// Примечание
				elemCont.Line.Add(GOST_21_110_2013.ElementNote(elemCont));
			}
		}

	} // class GOST_P_21_101_2020_Table_8

	class GOST_P_21_101_2020_Title_12
	{
		public static void FillTitle(int a)
		{
			//a.ToString();
			// Get data:
			string OrgName = Rvt.Handler.ProjInfo.OrganizationName;                                 // "_1"
			string FullOrgName = Rvt.Handler.ProjInfo.OrganizationDescription;                      // "_2"
			string ObjName = Rvt.Handler.ProjInfo.Name + "\r\n" + Rvt.Handler.ProjInfo.Address;     // "_4"
			string DocType = "Вид документации";                                                    // "_5"
			string DocName = "Наименование документа";                                              // "_6"
			string TomName = "Обозначение документа или тома";                                      // "_7"
			string TomNumber = "Номер тома";                                                        // "_8"
			string AuthorsPositions = "________________________\r\n(должность)";                    // "_9"
			string Date = "(подпись)\r\n";                                                          // "_10"

			Date += DateTime.UtcNow.ToString("dd.MM.yyyy");
			string AuthorsNames = Rvt.Handler.ProjInfo.Author;                                      // "_11"

			string Year = DateTime.UtcNow.ToString("yyyy");                                         // "_12"
			string Agreement = String.Empty;                                                        // "_3"
			//Agreement = "СОГЛАСОВАНО\r\n________________________\r\n" +
			//	"________________________\r\n«___» __________ " +
			//	Year + " г.";

			// Fill data
			Work.Book.WSs.Last().Tables[0].Data = new List<List<string>>() {
				new List<string>() { OrgName },
				new List<string>() { FullOrgName },
				new List<string>() { Agreement },
				new List<string>() { ObjName },
				new List<string>() { DocType },
				new List<string>() { DocName },
				new List<string>() { TomName },
				new List<string>() { TomNumber },
				new List<string>() { AuthorsPositions, Date, AuthorsNames },
				new List<string>() { Year }
			};
		}

	} // class GOST_P_21_101_2020_Title_12

} // namespace RevitToGOST
