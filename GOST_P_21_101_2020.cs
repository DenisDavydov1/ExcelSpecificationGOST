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

namespace RevToGOSTv0
{
	class GOST_P_21_101_2020_Title_12 : IGostData
	{
		/*
		** Member fields
		*/

		// наименование вышестоящей организации "_1"
		private string OrgName;

		// логотип (не обязательно), полное наименование организации, подготовившей документ "_2"
		private string FullOrgName;

		// в левой части — гриф согласования, в правой части — гриф утверждения,
		// выполняемые по ГОСТ Р 7.0.97 (при необходимости) "_3"
		private string Agreement;

		// наименование объекта капитального строительства, этап и вид строительства (при необходимости) "_4"
		private string ObjName;

		// вид документации (при необходимости) "_5"
		private string DocType;

		// наименование документа "_6"
		private string DocName;

		// обозначение документа или тома "_7"
		private string TomName;

		// номер тома "_8"
		private string TomNumber;

		// должности лиц, ответственных за разработку документа "_9"
		private string AuthorsPositions;

		// подписи лиц, указанных на поле 9, и даты подписания или отметки об ЭП,
		// выполняемые согласно ГОСТ Р 7.0.97;"_10"
		private string Date;

		// инициалы и фамилии лиц, указанных на поле 9 "_11"
		private string AuthorsNames;

		// год выпуска документа "_12"
		private string Year;

		/*
		** Member methods
		*/

		public void FillLines()
		{
			// "_1"
			OrgName = Rvt.Handler.ProjInfo.OrganizationName;
			
			// "_2"
			FullOrgName = Rvt.Handler.ProjInfo.OrganizationDescription;
			
			// "_4"
			ObjName = Rvt.Handler.ProjInfo.Name + "\r\n" + Rvt.Handler.ProjInfo.Address;

			// "_5"
			DocType = "Вид документации";

			// "_6"
			DocName = "Наименование документа";

			// "_7"
			TomName = "Обозначение документа или тома";

			// "_8"
			TomNumber = "Номер тома";

			// "_9"
			AuthorsPositions = "________________________\r\n(должность)";

			// "_10"
			Date = "(подпись)\r\n";
			Date += DateTime.UtcNow.ToString("dd.MM.yyyy");

			// "_11"
			AuthorsNames = Rvt.Handler.ProjInfo.Author;

			// "_12"
			Year = DateTime.UtcNow.ToString("yyyy");

			// "_3"
			Agreement = String.Empty;
			//Agreement = "СОГЛАСОВАНО\r\n________________________\r\n" +
			//	"________________________\r\n«___» __________ " +
			//	Year + " г.";
		}

		public List<List<string>> FillList()
		{
			List<List<string>> output = new List<List<string>>()
			{
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
			return output;
		}
	} // class GOST_P_21_101_2020_Title_12

} // namespace RevToGOSTv0
