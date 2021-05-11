using System;
using System.Collections.Generic;

namespace RevitToGOST
{
	class GOST_P_2_106_2019_Table_1
	{
		public static void FillLine(ElementContainer elemCont)
		{
			//// Заголовок категории
			if (elemCont.LineType == ElementContainer.ContType.Category)
			{
				elemCont.Line = new List<string>() { "", "", "", "", elemCont.CategoryLine, "", "" };
			}

			//// Нумерация стоблцов
			else if (elemCont.LineType == ElementContainer.ContType.ColumnsEnumeration)
			{
				elemCont.Line = new List<string>() { "1", "2", "3", "4", "5", "6", "7" };
			}

			//// Настоящий элемент
			else
			{
				elemCont.Line = new List<string>();

				//(*1): Формат
				elemCont.Line.Add("");

				//(*2): Зона
				elemCont.Line.Add("");

				//(*3): Поз.
				elemCont.Line.Add(elemCont.Position.ToString());

				//(*4): Обозначение
				elemCont.Line.Add(GOST_21_110_2013.ElementType(elemCont));

				//(*5): Наименование
				elemCont.Line.Add(GOST_21_110_2013.ElementName(elemCont));

				//(*6): Кол.
				double amount = GOST_21_110_2013.ElementAmount(elemCont);
				if (amount == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(amount.ToString());

				//(*7): Примечаение
				elemCont.Line.Add(GOST_21_110_2013.ElementNote(elemCont));
			}
		}
	}

	class GOST_P_2_106_2019_Table_5
	{
		public static void FillLine(ElementContainer elemCont)
		{
			//// Заголовок категории
			if (elemCont.LineType == ElementContainer.ContType.Category)
			{
				elemCont.Line = new List<string>() { "", elemCont.CategoryLine, "", "", "", "", "", "", "", "", "" };
			}

			//// Нумерация стоблцов
			else if (elemCont.LineType == ElementContainer.ContType.ColumnsEnumeration)
			{
				elemCont.Line = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11" };
			}

			//// Настоящий элемент
			else
			{
				elemCont.Line = new List<string>();

				//(*1): № строки
				elemCont.Line.Add(elemCont.Position.ToString());

				//(*2): Наименование
				elemCont.Line.Add(GOST_21_110_2013.ElementName(elemCont));
				// ??? elemCont.Line.Add(GOST_21_110_2013.ElementType(elemCont));

				//(*3): Код продукции
				elemCont.Line.Add(GOST_21_110_2013.ElementProdCode(elemCont));

				//(*4): Обозначение документа на поставку
				elemCont.Line.Add("");

				//(*5): Поставщик
				elemCont.Line.Add(GOST_21_110_2013.ElementProvider(elemCont));

				//(*6): Куда входит(обозначение)
				elemCont.Line.Add("");

				//(*7): Количество
				elemCont.Line.Add("");
				elemCont.Line.Add("");
				elemCont.Line.Add("");

				double amount = GOST_21_110_2013.ElementAmount(elemCont);
				if (amount == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(amount.ToString());

				//(*8): Примечание
				elemCont.Line.Add(GOST_21_110_2013.ElementNote(elemCont));
			}
		}
	}
}
