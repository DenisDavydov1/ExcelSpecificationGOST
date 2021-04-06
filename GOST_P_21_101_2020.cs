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
	class GOST_P_21_101_2020_Title_12 : IGostData
	{
		public static void FillLine(ElementContainer elemCont)
		{
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
