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
	[TransactionAttribute(TransactionMode.Manual)]
	[RegenerationAttribute(RegenerationOption.Manual)]
	public class MainClass : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			Log.ClearLog();
			Work.Book = new WorkBook();

			Rvt.Handler = new RvtHandler(commandData, elements);
			Rvt.Data = new RvtData();
			Rvt.Control = new RvtControl();

			MainWindow mainWin = new MainWindow();
			mainWin.Show();

			return Result.Succeeded;
		}
	} // class MainClass

} // namespace RevitToGOST
