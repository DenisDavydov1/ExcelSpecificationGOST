using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitToGOST
{
	class GOST_21_301_2014_Title_2
	{
		public static void FillTitle(int a)
		{
			string _1 = Rvt.Handler.ProjInfo.OrganizationDescription;//(1): Наименование вышестоящей организации
			string _2 = Rvt.Handler.ProjInfo.OrganizationName + "\r\n" + Rvt.Handler.ProjInfo.Address;//(2): Логотип / наименование организации, подготовившей документацию
			string _3 = "";//(3): Номер и дата выдачи документа
			string _4 = Rvt.Handler.ProjInfo.ClientName;//(4): Наименование организации заказчика
			string _5 = "";//(5): Гриф ограничение доступа
			string _6 = Rvt.Handler.ProjInfo.BuildingName;//(6): Наименование объекта капитального строительства
			string _7 = Rvt.Handler.ProjInfo.Name;//(7): Наименование отчета
			string _8 = Rvt.Handler.ProjInfo.Address;//(8): Обозначение отчета
			string _9 = "";//(9): Номер тома
			string _10 = "________________________\r\n(должность)";//(10): Должности лиц, ответственных за разработку отчета
			string _11 = "(подпись)\r\n" + DateTime.UtcNow.ToString("dd.MM.yyyy");//(11): Подписи лиц, ответственных за разработку отчета
			string _12 = Rvt.Handler.ProjInfo.Author;//(12): Инициалы и фамилии лиц, указанных в поле 11
			string _13 = DateTime.UtcNow.ToString("yyyy");//(13): Место и год выпуска
			string _14 = "";//(14): Размещение таблицы регистрации изменений

			Work.Book.WSs.Last().Tables[0].Data = new List<List<string>>() {
				new List<string>() { _1 },
				new List<string>() { _2 },
				new List<string>() { _3 },
				new List<string>() { _4, _5 },
				new List<string>() { _6 },
				new List<string>() { _7 },

				new List<string>() { _8 },
				new List<string>() { _9 },

				new List<string>() { _10, _11, _12 },
				new List<string>() { _14, _13 }
			};
		}
	}
}
