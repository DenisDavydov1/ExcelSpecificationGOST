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
using System.ComponentModel;

namespace RevitToGOST
{
	[TransactionAttribute(TransactionMode.Manual)]
	[RegenerationAttribute(RegenerationOption.Manual)]
	public class MainClass : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			try
			{
				Log.ClearLog();

				Rvt.Windows = new RvtWindows();
				Rvt.Windows.RunLoadingWindow();

				Work.Book = new WorkBook();
				Rvt.Handler = new RvtHandler(commandData, elements);
				Rvt.Control = new RvtControl();
				Rvt.Data = new RvtData();

				Rvt.Windows.CloseLoadingWindow();
				Rvt.Windows.RunMainWindow();
			}
			catch (Exception e) { Log.WriteLine("Caught exception from mainclass:\n{0}\n{1}", e.Message, e.StackTrace); }

			return Result.Succeeded;
		}

	} // class MainClass

} // namespace RevitToGOST
