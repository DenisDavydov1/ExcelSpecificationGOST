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
using System.Threading;
//using System.Reflection;

namespace RevitToGOST // RevitToExcelGOST
{
	[Transaction(TransactionMode.Manual)]
	[Regeneration(RegenerationOption.Manual)]
	[Journaling(JournalingMode.NoCommandData)]
	public class RevitToGOSTApp : IExternalApplication
	{
		static private readonly string RibbonPanelName = "Revit в Excel по ГОСТ";
		static private readonly string AddInPath = typeof(RevitToGOSTApp).Assembly.Location;

		#region IExternalApplication Members
		public Result OnShutdown(UIControlledApplication application)
		{
			// remove events??
			return Result.Succeeded;
		}

		public Result OnStartup(UIControlledApplication application)
		{
			try
			{
				CreateRibbonPanel(application);
				return Result.Succeeded;
			}
			catch (Exception ex)
			{
				TaskDialog.Show("Error", ex.ToString());
				return Result.Failed;
			}
		}
		#endregion

		private void CreateRibbonPanel(UIControlledApplication application)
		{
			// Add a new ribbon panel to Revit Add-Ins tab
			RibbonPanel ribbonPanel = application.CreateRibbonPanel(RibbonPanelName);
			PushButtonData pushButtonData = new PushButtonData(
				"RvtToGst",
				"RvtToGst",
				AddInPath,
				"RevitToGOST.RevitToExcelGOSTCommand");
			//pushButtonData.Image = Bitmaps.Convert(Properties.Resources.ButtonIcon);
			PushButton pushButton = ribbonPanel.AddItem(pushButtonData) as PushButton;
			pushButton.ToolTip = "Экспортировать модель в спецификацию по ГОСТ в Excel";
			pushButton.LargeImage = Bitmaps.Convert(Properties.Resources.ButtonIcon);

			//application.ControlledApplication.DocumentCreated += 

			//string location = Assembly.GetExecutingAssembly().Location;
			//var pushButtonData = new PushButtonData(
			//	"MyApp_Browser_Command",
			//	"MyApp",
			//	location,
			//	"MyApp.ShowWindow");
			//pushButtonData.AvailabilityClassName = "MyApp.ShowWindow";
			//PushButton button = ribbonPanel.AddItem(pushButtonData) as PushButton;
			//application.ControlledApplication.ApplicationInitialized += ControlledApplication_ApplicationInitialized;
		}
	}

	[Transaction(TransactionMode.Manual)]
	[Regeneration(RegenerationOption.Manual)]
	public class RevitToExcelGOSTCommand : IExternalCommand
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
				Work.Bitmaps = new Bitmaps();

				Rvt.Windows.CloseLoadingWindow();
				Rvt.Windows.RunMainWindow();
			}
			catch (Exception e)
			{
				message = e.Message;
				Rvt.Handler.Result = Result.Failed;
				Log.WriteLine("Caught exception from mainclass:\n{0}\n{1}", e.Message, e.StackTrace);
			}
			return Rvt.Handler == null ? Result.Failed : Rvt.Handler.Result;
		}

	} // class RevitToExcelGOSTCommand



	//[TransactionAttribute(TransactionMode.Manual)]
	//[RegenerationAttribute(RegenerationOption.Manual)]
	//public class MainClass : IExternalCommand
	//{
	//	public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
	//	{
	//		try
	//		{
	//			Log.ClearLog();

	//			Rvt.Windows = new RvtWindows();
	//			Rvt.Windows.RunLoadingWindow();

	//			Work.Book = new WorkBook();
	//			Rvt.Handler = new RvtHandler(commandData, elements);
	//			Rvt.Control = new RvtControl();
	//			Rvt.Data = new RvtData();
	//			Work.Bitmaps = new Bitmaps();

	//			Rvt.Windows.CloseLoadingWindow();
	//			Rvt.Windows.RunMainWindow();
	//		}
	//		catch (Exception e)
	//		{
	//			message = e.Message;
	//			Rvt.Handler.Result = Result.Failed;
	//			Log.WriteLine("Caught exception from mainclass:\n{0}\n{1}", e.Message, e.StackTrace);
	//		}
	//		return Rvt.Handler == null ? Result.Failed : Rvt.Handler.Result;
	//	}

	//} // class MainClass

} // namespace RevitToGOST
