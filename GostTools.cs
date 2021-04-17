﻿using System;
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
using System.Windows.Media.Imaging;

using RevitToGOST.Properties;
using System.Drawing;
using System.Windows.Data;
using System.IO;

namespace RevitToGOST
{
	static class GostTools
	{
		public static ElementSet ElementConvert(IList<Element> list)
		{
			ElementSet elemSet = new ElementSet();
			foreach (Element elem in list)
				elemSet.Insert(elem);
			return elemSet;
		}

	} // class GostTools

	static class Constants
	{
		public const string LogPath = @"F:\CS_CODE\REVIT\PROJECTS\tmp";
		public const string DefaultFilePath = @"F:\CS_CODE\REVIT\PROJECTS\output.xlsx";
		public const double inch = 2.54;
		public const double mm_w = 0.483;
		public const double mm_h = 2.9;
		public const string DefaultName = "Без названия";

		public struct A3
		{
			public const int Height = 420;
			public const int Width = 297;
		}

		public struct A4
		{
			public const int Height = 297;
			public const int Width = 210;
		}

	} // static class Constants

	static class ConfFile
	{
		public static readonly string[] Conf = {
			String.Empty,
			Encoding.UTF8.GetString(Resources.GOST_21_110_2013_Table_1),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Stamp_3),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Dop_3),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Stamp_4),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Dop_4),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Stamp_5),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Dop_5),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Stamp_6),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Dop_6),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Table_7),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Table_8),
			String.Empty,	// misc 9
			String.Empty,	// misc 9a
			String.Empty,	// misc 10
			String.Empty,	// misc 11
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Title_12),
			Encoding.UTF8.GetString(Resources.GOST_P_21_101_2020_Title_12a),
			String.Empty,	// title 14
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Stamp_1),
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Dop_1),
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Stamp_2),
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Dop_2),
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Stamp_2a),
			Encoding.UTF8.GetString(Resources.GOST_2_104_2006_Dop_2a),
			Encoding.UTF8.GetString(Resources.GOST_21_301_2014_Title_2),
			Encoding.UTF8.GetString(Resources.GOST_P_2_106_2019_Table_1),
			Encoding.UTF8.GetString(Resources.GOST_P_2_106_2019_Table_5)
		};

		public static readonly int[] Lines = {
			0,	// None
			24,	// GOST_21_110_2013_Table1
			0,	// GOST_P_21_101_2020_Stamp3
			0,	// GOST_P_21_101_2020_Dop3
			0,	// stamp 4
			0,	// dop 4
			0,	// stamp 5
			0,	// dop 5
			0,	// stamp 6
			0,	// dop 6
			26,	// table 7
			41,	// table 8
			0,	// misc 9
			0,	// misc 9a
			0,	// misc 10
			0,	// misc 11
			0,	// GOST_P_21_101_2020_Title_12
			0,	// GOST_P_21_101_2020_Title_12a
			0,	// title 14
			0,	// GOST_2_104_2006_Stamp_1
			0,	// GOST_2_104_2006_Dop_1
			0,	// GOST_2_104_2006_Stamp_2
			0,	// GOST_2_104_2006_Dop_2
			0,	// GOST_2_104_2006_Stamp_2a
			0,	// GOST_2_104_2006_Dop_2a
			0,	// GOST_21_301_2014_Title_2
			26, // GOST_P_2_106_2019_Table_1
			24	// GOST_P_2_106_2019_Table_5
		};

		public static readonly Action<int>[] FillTitle = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			null,	// GOST_P_21_101_2020_Stamp3
			null,	// GOST_P_21_101_2020_Dop3
			null,	// stamp 4
			null,	// dop 4
			null,	// stamp 5
			null,	// dop 5
			null,	// stamp 6
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			GOST_P_21_101_2020_Title_12.FillTitle,
			GOST_P_21_101_2020_Title_12a.FillTitle,
			null,	// title 14
			null,   // GOST_2_104_2006_Stamp_1
			null,	// GOST_2_104_2006_Dop_1
			null,   // GOST_2_104_2006_Stamp_2
			null,	// GOST_2_104_2006_Dop_2
			null,   // GOST_2_104_2006_Stamp_2a
			null,	// GOST_2_104_2006_Dop_2a
			GOST_21_301_2014_Title_2.FillTitle,
			null,	//GOST_P_2_106_2019_Table_1
			null	//GOST_P_2_106_2019_Table_5
		};

		public static readonly Action<ElementContainer>[] FillLine = {
			null,	// None
			GOST_21_110_2013.FillLine,
			null,	// GOST_P_21_101_2020_Stamp3
			null,	// GOST_P_21_101_2020_Dop3
			null,	// stamp 4
			null,	// dop 4
			null,	// stamp 5
			null,	// dop 5
			null,	// stamp 6
			null,	// dop 6
			GOST_P_21_101_2020_Table_7.FillLine,
			GOST_P_21_101_2020_Table_8.FillLine,
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// GOST_P_21_101_2020_Title_12a
			null,	// title 14
			null,   // GOST_2_104_2006_Stamp_1
			null,	// GOST_2_104_2006_Dop_1
			null,   // GOST_2_104_2006_Stamp_2
			null,	// GOST_2_104_2006_Dop_2
			null,   // GOST_2_104_2006_Stamp_2a
			null,	// GOST_2_104_2006_Dop_2a
			null,	// GOST_21_301_2014_Title_2
			GOST_P_2_106_2019_Table_1.FillLine,
			GOST_P_2_106_2019_Table_5.FillLine,
		};

		public static readonly Action<int>[] FillStamp = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			GOST_P_21_101_2020_Stamp_3.FillStamp,
			null,	// GOST_P_21_101_2020_Dop3
			GOST_P_21_101_2020_Stamp_4.FillStamp,
			null,	// dop 4
			GOST_P_21_101_2020_Stamp_5.FillStamp,
			null,	// dop 5
			GOST_P_21_101_2020_Stamp_6.FillStamp,
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// GOST_P_21_101_2020_Title_12a
			null,	// title 14
			GOST_2_104_2006_Stamp_1.FillStamp,
			null,	// GOST_2_104_2006_Dop_1
			GOST_2_104_2006_Stamp_2.FillStamp,
			null,	// GOST_2_104_2006_Dop_2
			GOST_2_104_2006_Stamp_2a.FillStamp,
			null,	// GOST_2_104_2006_Dop_2a
			null,	// GOST_21_301_2014_Title_2
			null,	//GOST_P_2_106_2019_Table_1
			null	//GOST_P_2_106_2019_Table_5
		};

		public static readonly Action<int>[] FillDop = {
			null,	// None
			null,	// GOST_21_110_2013_Table1
			null,	// GOST_P_21_101_2020_Stamp3
			GOST_P_21_101_2020_Dop_3.FillDop,
			null,	// stamp 4
			GOST_P_21_101_2020_Dop_4.FillDop,
			null,	// stamp 5
			GOST_P_21_101_2020_Dop_5.FillDop,
			null,	// stamp 6
			GOST_P_21_101_2020_Dop_6.FillDop,
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			null,	// GOST_P_21_101_2020_Title_12
			null,	// GOST_P_21_101_2020_Title_12a
			null,	// title 14
			null,	// GOST_2_104_2006_Stamp_1
			GOST_2_104_2006_Dop_1.FillDop,
			null,	// GOST_2_104_2006_Stamp_2
			GOST_2_104_2006_Dop_2.FillDop,
			null,	// GOST_2_104_2006_Stamp_2a
			GOST_2_104_2006_Dop_2a.FillDop,
			null,	// GOST_21_301_2014_Title_2
			null,	//GOST_P_2_106_2019_Table_1
			null	//GOST_P_2_106_2019_Table_5
		};

		/*
			// None
			// GOST_21_110_2013_Table_1
			// GOST_P_21_101_2020_Stamp_3
			// GOST_P_21_101_2020_Dop_3
			// GOST_P_21_101_2020_Stamp_4
			// GOST_P_21_101_2020_Dop_4
			// GOST_P_21_101_2020_Stamp_5
			// GOST_P_21_101_2020_Dop_5
			// GOST_P_21_101_2020_Stamp_6
			// GOST_P_21_101_2020_Dop_6
			// GOST_P_21_101_2020_Table_7
			// GOST_P_21_101_2020_Table_8
			// GOST_P_21_101_2020_Misc_9
			// GOST_P_21_101_2020_Misc_9a
			// GOST_P_21_101_2020_Misc_10
			// GOST_P_21_101_2020_Misc_11
			// GOST_P_21_101_2020_Title_12
			// GOST_P_21_101_2020_Title_12a
			// GOST_P_21_101_2020_Title_14
			// GOST_2_104_2006_Stamp_1
			// GOST_2_104_2006_Dop_1
			// GOST_2_104_2006_Stamp_2
			// GOST_2_104_2006_Dop_2
			// GOST_2_104_2006_Stamp_2a
			// GOST_2_104_2006_Dop_2a
			// GOST_21_301_2014_Title_2
			// GOST_P_2_106_2019_Table_1
			// GOST_P_2_106_2019_Table_5
		*/

		public static readonly string[] Descriprions = {
			String.Empty, // None
			/* GOST_21_110_2013_Table_1 */ "(*1): Поз.\n(*2): Наименование и техническая характеристика\n(*3): Тип, марка, обозначение документа, опросного листа\n(*4): Код продукции\n(*5): Поставщик\n(*6): Ед. измерения\n(*7): Количество\n(*8): Масса ед., кг\n(*9): Примечание\n", // GOST_21_110_2013_Table_1
			/* GOST_P_21_101_2020_Stamp_3 */ "(1): Обозначение документа\n(2): Наименование предприятия; объекта строительства\n(3) : Наименование здания(сооружения)\n(4) : Наименование изображений на листе\n(6) : Условное обозначение вида документации\n(7) : Порядковый номер листа документа\n(8) : Общее кол-во листов документа\n(9) : Наименование организации, разработавшей документ\n(10) : Характер работы лица, подписывающего документ\n(11) : Фамилии лиц, подписывающих документ\n(12) : Подписи лиц, фамилии которых указаны в графе 11\n(13) : Дата подписания документа\n(14-19) : Сведения об изменениях\n", // GOST_P_21_101_2020_Stamp_3
			/* GOST_P_21_101_2020_Dop_3 */ "(10): Характер работы лица, подписывающего документ/n(11): Фамилии лиц, подписывающих документ/n(12): Подписи лиц, фамилии которых указаны в графе 11/n(13): Дата подписания документа/n(20): Инвентарный номер подлинника/n(21): Подпись лица, принявшего подлинник на хранение, и дата приемки/n(22): Инвентарный номер подлинника, взамен которого выпущен документ/n", // GOST_P_21_101_2020_Dop_3
			/* GOST_P_21_101_2020_Stamp_4 */ "(1): Обозначение документа\n(5): Наименование изделия/документа\n(6): Условное обозначение вида документации\n(7): Порядковый номер листа документа\n(8): Общее кол-во листов документа\n(9): Наименование организации, разработавшей документ\n(10): Характер работы лица, подписывающего документ\n(11): Фамилии лиц, подписывающих документ\n(12): Подписи лиц, фамилии которых указаны в графе 11\n(13): Дата подписания документа\n(14-19): Сведения об изменениях\n(23): Обозначение материала изделия\n(24): Масса изделия\n(25): Масштаб\n(27): Краткое наименование организации-заказчика\n", // GOST_P_21_101_2020_Stamp_4
			/* GOST_P_21_101_2020_Dop_4 */ "(20): Инвентарный номер подлинника\n(21): Подпись лица, принявшего подлинник на хранение, и дата приемки\n(22): Инвентарный номер подлинника, взамен которого выпущен документ\n", // GOST_P_21_101_2020_Dop_4
			/* GOST_P_21_101_2020_Stamp_5 */ "(1): Обозначение документа\n(5): Наименование изделия/документа\n(6): Условное обозначение вида документации\n(7): Порядковый номер листа документа\n(8): Общее кол-во листов документа\n(9): Наименование организации, разработавшей документ\n(10): Характер работы лица, подписывающего документ\n(11): Фамилии лиц, подписывающих документ\n(12): Подписи лиц, фамилии которых указаны в графе 11\n(13): Дата подписания документа\n(14-19): Сведения об изменениях\n(27): Краткое наименование организации-заказчика\n", // GOST_P_21_101_2020_Stamp_5
			/* GOST_P_21_101_2020_Dop_5 */ "(10): Характер работы лица, подписывающего документ\n(11): Фамилии лиц, подписывающих документ\n(12): Подписи лиц, фамилии которых указаны в графе 11\n(13): Дата подписания документа\n(20): Инвентарный номер подлинника\n(21): Подпись лица, принявшего подлинник на хранение, и дата приемки\n(22): Инвентарный номер подлинника, взамен которого выпущен документ\n", // GOST_P_21_101_2020_Dop_5
			/* GOST_P_21_101_2020_Stamp_6 */ "(1): Обозначение документа\n(7): Порядковый номер листа документа\n(14-19): Сведения об изменениях\n", // GOST_P_21_101_2020_Stamp_6
			/* GOST_P_21_101_2020_Dop_6 */ "(20): Инвентарный номер подлинника\n(21): Подпись лица, принявшего подлинник на хранение, и дата приемки\n(22): Инвентарный номер подлинника, взамен которого выпущен документ\n", // GOST_P_21_101_2020_Dop_6
			/* GOST_P_21_101_2020_Table_7 */ "(*1): Поз.\n(*2): Обозначение\n(*3): Наименование\n(*4): Кол.\n(*5): Масса ед., кг\n(*6): Примечание\n", // GOST_P_21_101_2020_Table_7
			/*GOST_P_21_101_2020_Table_8  */ "(*1): Поз.\n(*2): Обозначение\n(*3): Наименование\n(*4): Кол.\n(*5): Масса ед., кг\n(*6): Примечание\n", // GOST_P_21_101_2020_Table_8
			/*  */ String.Empty, // GOST_P_21_101_2020_Misc_9
			/*  */ String.Empty, // GOST_P_21_101_2020_Misc_9a
			/*  */ String.Empty, // GOST_P_21_101_2020_Misc_10
			/*  */ String.Empty, // GOST_P_21_101_2020_Misc_11
			/* GOST_P_21_101_2020_Title_12 */ "(1): Наименование вышестоящей организации\n(2): Логотип/полное наименование организации, подготовившей документ\n(3): Гриф согласования, гриф утверждения\n(4): Наименование объекта капитального строительства\n(5): Вид документации\n(6): Наименование документа\n(7): Обозначение документа или тома\n(8): Номер тома\n(9): Должности лиц, ответственных за разработку документа\n(10): Подписи лиц, указанных на поле 9 и даты подписания\n(11): Инициалы и фамилии лиц, указанных на поле 9\n(12): Год выпуска документа\n", // GOST_P_21_101_2020_Title_12
			/* GOST_P_21_101_2020_Title_12a */ "(1): Наименование вышестоящей организации\n(2): Логотип/полное наименование организации, подготовившей документ\n(3): Гриф согласования, гриф утверждения\n(4): Наименование объекта капитального строительства\n(5): Вид документации\n(6): Наименование документа\n(7): Обозначение документа или тома\n(8): Номер тома\n(9): Должности лиц, ответственных за разработку документа\n(10): Подписи лиц, указанных на поле 9 и даты подписания\n(11): Инициалы и фамилии лиц, указанных на поле 9\n(12): Год выпуска документа\n", // GOST_P_21_101_2020_Title_12a
			/*  */ String.Empty,	// GOST_P_21_101_2020_Title_14
			/* GOST_2_104_2006_Stamp_1 */ "(1): Наименование изделия\n(2): Обозначение документа\n(3): Обозначение материала\n(4): Литера документа\n(5): Масса изделия\n(6): Масштаб\n(7): Порядковый номер листа\n(8): Обще кол-во листов документа\n(9): Наименование организации\n(10): Характер работы лица, подписывающего документ\n(11): Фамилии лиц, подписывающих документ\n(12): Подписи лиц, фамилии которых указаны в графе 11\n(13): Дата подписания документа\n(14-18): Сведения об изменениях\n(30): Индекс заказчика\n", // GOST_2_104_2006_Stamp_1
			/* GOST_2_104_2006_Dop_1 */ "(19): Инвентарный номер подлинника\n(20): Сведения о приемке дубликата\n(21): Инвентарный номер подлинника, взамен которого выпущен данный документ\n(22): Инвентарный номер дубликата\n(23): Сведения о приеме дубликата в службу\n",	// GOST_2_104_2006_Dop_1
			/* GOST_2_104_2006_Stamp_2 */ "(1): Наименование изделия\n(2): Обозначение документа\n(4): Литера документа\n(7): Порядковый номер листа\n(8): Обще кол-во листов документа\n(9): Наименование организации\n(10): Характер работы лица, подписывающего документ\n(11): Фамилии лиц, подписывающих документ\n(12): Подписи лиц, фамилии которых указаны в графе 11\n(13): Дата подписания документа\n(14-18): Сведения об изменениях\n(27): Знак, устанавливаемый заказчиком\n(28): Номер решения и год утверждения документации литеры\n(29): Номер решения и год утверждения документации\n(30): Индекс заказчика\n", // GOST_2_104_2006_Stamp_2
			/* GOST_2_104_2006_Dop_2 */ "(19): Инвентарный номер подлинника\n(20): Сведения о приемке дубликата\n(21): Инвентарный номер подлинника, взамен которого выпущен данный документ\n(22): Инвентарный номер дубликата\n(23): Сведения о приеме дубликата в службу\n",	// GOST_2_104_2006_Dop_2
			/* GOST_2_104_2006_Stamp_2a */ "(2): Обозначение документа\n(7): Порядковый номер листа\n(14-18): Сведения об изменениях\n", // GOST_2_104_2006_Stamp_2a
			/* GOST_2_104_2006_Dop_2a */ "(19): Инвентарный номер подлинника\n(20): Сведения о приемке дубликата\n(21): Инвентарный номер подлинника, взамен которого выпущен данный документ\n(22): Инвентарный номер дубликата\n(23): Сведения о приеме дубликата в службу\n",	// GOST_2_104_2006_Dop_2a
			/* GOST_21_301_2014_Title_2 */ "(1): Наименование вышестоящей организации\n(2): Логотип/наименование организации, подготовившей документацию\n(3): Номер и дата выдачи документа\n(4): Наименование организации заказчика\n(5): Гриф ограничение доступа\n(6): Наименование объекта капитального строительства\n(7): Наименование отчета\n(8): Обозначение отчета\n(9): Номер тома\n(10): Должности лиц, ответственных за разработку отчета\n(11): Подписи лиц, ответственных за разработку отчета\n(12): Инициалы и фамилии лиц, указанных в поле 11\n(13): Место и год выпуска\n(14): Размещение таблицы регистрации изменений\n", // GOST_21_301_2014_Title_2
			/* GOST_P_2_106_2019_Table_1 */ "(*1): Формат\n(*2): Зона\n(*3): Поз.\n(*4): Обозначение\n(*5): Наименование\n(*6): Кол.\n(*7): Примечаение\n",	// GOST_P_2_106_2019_Table_1
			/* GOST_P_2_106_2019_Table_5 */ "(*1): № строки\n(*2): Наименование\n(*3): Код продукции\n(*4): Обозначение документа на поставку\n(*5): Поставщик\n(*6): Куда входит (обозначение)\n(*7): Количество\n(*8): Примечание\n"	// GOST_P_2_106_2019_Table_5
		};

	} // class ConfFile

	static class ClassExtensions
	{
		public static int TotalLinesCount(this Dictionary<string, ElementCollection> dict)
		{
			int lines = dict.Keys.Count;
			foreach (ElementCollection elemCol in dict.Values)
				lines += elemCol.Count;
			return lines;
		}
	}

} // namespace RevitToGOST
