using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitToGOST
{
	class GOST_2_104_2006_Stamp_1
	{
		public static void FillStamp(int pageNumber)
		{
			string _1 = Rvt.Handler.ProjInfo.BuildingName;//(1): Наименование изделия
			string _2 = Rvt.Handler.ProjInfo.Name;//(2): Обозначение документа
			string _3 = "";//(3): Обозначение материала
			string _4 = "";//(4): Литера документа
			string _5 = "";//(5): Масса изделия
			string _6 = "";//(6): Масштаб
			string _7 = (pageNumber + 1).ToString();//(7): Порядковый номер листа
			string _8 = Work.Book.Pages.ToString();//(8): Обще кол-во листов документа
			string _9 = Rvt.Handler.ProjInfo.OrganizationName;//(9): Наименование организации
			string _10 = "";//(10): Характер работы лица, подписывающего документ
			string _11 = "";//(11): Фамилии лиц, подписывающих документ
			string _12 = "";//(12): Подписи лиц, фамилии которых указаны в графе 11
			string _13 = DateTime.UtcNow.ToString("MM.yy");//(13): Дата подписания документа
			//(14 - 18): Сведения об изменениях
			string _30 = Rvt.Handler.ProjInfo.ClientName;//(30): Индекс заказчика

			string AuthorName = Rvt.Handler.ProjInfo.Author;    // Разраб.

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "Изм.", "Лист", "№ докум.", "Подп.", "Дата" },

				new List<string>() { "Разраб.", AuthorName, "", _13 },
				new List<string>() { "Пров.", "", "", "" },
				new List<string>() { "Т.контр.", "", "", "" },
				new List<string>() { _10, _11, _12, "" },
				new List<string>() { "Н.контр.", "", "", "" },
				new List<string>() { "Утв.", "", "", "" },

				new List<string>() { _30, _2, _1, _3, _9 },
				new List<string>() { "Лит.", "Масса", "Масштаб" },
				new List<string>() { _4, "", "", _5, _6 },
				new List<string>() { "Лист " + _7, "Листов " + _8 }
			};
		}

	} // class GOST_2_104_2006_Stamp_1

	class GOST_2_104_2006_Dop_1
	{
		public static void FillDop(int pageNumber)
		{
			string _19 = "";//(19): Инвентарный номер подлинника
			string _20 = "";//(20): Сведения о приемке дубликата
			string _21 = "";//(21): Инвентарный номер подлинника, взамен которого выпущен данный документ
			string _22 = "";//(22): Инвентарный номер дубликата
			string _23 = "";//(23): Сведения о приеме дубликата в службу

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { "Подп. и дата", _23 },
				new List<string>() { "Инв. № дубл.", _22 },
				new List<string>() { "Взам. инв. №", _21 },
				new List<string>() { "Подп. и дата", _20 },
				new List<string>() { "Инв. № подл.", _19 }
			};
		}

	} // class GOST_2_104_2006_Dop_1

	class GOST_2_104_2006_Stamp_2
	{
		public static void FillStamp(int pageNumber)
		{
			string _1 = Rvt.Handler.ProjInfo.BuildingName;//(1): Наименование изделия
			string _2 = Rvt.Handler.ProjInfo.Name;//(2): Обозначение документа
			string _4 = "";//(4): Литера документа
			string _7 = (pageNumber + 1).ToString();//(7): Порядковый номер листа
			string _8 = Work.Book.Pages.ToString();//(8): Обще кол-во листов документа
			string _9 = Rvt.Handler.ProjInfo.OrganizationName;//(9): Наименование организации
			string _10 = "";//(10): Характер работы лица, подписывающего документ
			string _11 = "";//(11): Фамилии лиц, подписывающих документ
			string _12 = "";//(12): Подписи лиц, фамилии которых указаны в графе 11
			string _13 = DateTime.UtcNow.ToString("MM.yy");//(13): Дата подписания документа
			//(14 - 18): Сведения об изменениях
			string _27 = "";//(27): Знак, устанавливаемый заказчиком
			string _28 = "";//(28): Номер решения и год утверждения документации литеры
			string _29 = "";//(29): Номер решения и год утверждения документации
			string _30 = Rvt.Handler.ProjInfo.ClientName;//(30): Индекс заказчика

			string AuthorName = Rvt.Handler.ProjInfo.Author;    // Разраб.

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "Изм.", "Лист", "№ докум.", "Подп.", "Дата" },

				new List<string>() { "Разраб.", AuthorName, "", _13 },
				new List<string>() { "Пров.", "", "", "" },
				new List<string>() { _10, _11, _12, "" },
				new List<string>() { "Н.контр.", "", "", "" },
				new List<string>() { "Утв.", "", "", "" },

				new List<string>() { _27, _28, _29, _30 },
				new List<string>() { _2, _1, _9 },
				new List<string>() { "Лит.", "Лист", "Листов" },
				new List<string>() { _4, "", "", _7, _8 }
			};
		}

	} // class GOST_2_104_2006_Stamp_2

	class GOST_2_104_2006_Dop_2
	{
		public static void FillDop(int pageNumber)
		{
			string _19 = "";//(19): Инвентарный номер подлинника
			string _20 = "";//(20): Сведения о приемке дубликата
			string _21 = "";//(21): Инвентарный номер подлинника, взамен которого выпущен данный документ
			string _22 = "";//(22): Инвентарный номер дубликата
			string _23 = "";//(23): Сведения о приеме дубликата в службу

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { "Подп. и дата", _23 },
				new List<string>() { "Инв. № дубл.", _22 },
				new List<string>() { "Взам. инв. №", _21 },
				new List<string>() { "Подп. и дата", _20 },
				new List<string>() { "Инв. № подл.", _19 }
			};
		}

	} // class GOST_2_104_2006_Dop_2

	class GOST_2_104_2006_Stamp_2a
	{
		public static void FillStamp(int pageNumber)
		{
			string _2 = Rvt.Handler.ProjInfo.Name;//(2): Обозначение документа
			string _7 = (pageNumber + 1).ToString();//(7): Порядковый номер листа
			//(14 - 18): Сведения об изменениях

			Work.Book.WSs[pageNumber].Tables[1].Data = new List<List<string>>() {
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "", "", "", "", "" },
				new List<string>() { "Изм.", "Лист", "№ докум.", "Подп.", "Дата" },

				new List<string>() { _2 },
				new List<string>() { "Лист", _7 }
			};
		}

	} // class GOST_2_104_2006_Stamp_2a

	class GOST_2_104_2006_Dop_2a
	{
		public static void FillDop(int pageNumber)
		{
			string _19 = "";//(19): Инвентарный номер подлинника
			string _20 = "";//(20): Сведения о приемке дубликата
			string _21 = "";//(21): Инвентарный номер подлинника, взамен которого выпущен данный документ
			string _22 = "";//(22): Инвентарный номер дубликата
			string _23 = "";//(23): Сведения о приеме дубликата в службу

			Work.Book.WSs[pageNumber].Tables[2].Data = new List<List<string>>() {
				new List<string>() { "Подп. и дата", _23 },
				new List<string>() { "Инв. № дубл.", _22 },
				new List<string>() { "Взам. инв. №", _21 },
				new List<string>() { "Подп. и дата", _20 },
				new List<string>() { "Инв. № подл.", _19 }
			};
		}

	} // class GOST_2_104_2006_Dop_2a


} // namespace RevitToGOST
