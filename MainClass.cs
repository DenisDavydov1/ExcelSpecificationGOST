using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace RevitToGOST
{
	[Transaction(TransactionMode.Manual)]
	[Regeneration(RegenerationOption.Manual)]
	[Journaling(JournalingMode.NoCommandData)]
	public class RevitToGOSTApp : IExternalApplication
	{
		static private readonly string RibbonPanelName = "Revit по ГОСТ";
		static private readonly string AddInPath = typeof(RevitToGOSTApp).Assembly.Location;

		#region IExternalApplication Members
		public Result OnShutdown(UIControlledApplication application)
		{
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
			RibbonPanel ribbonPanel = application.CreateRibbonPanel(RibbonPanelName);
			PushButtonData pushButtonData = new PushButtonData(
				"Экспорт в Excel",
				"Экспорт в Excel",
				AddInPath,
				"RevitToGOST.RevitToExcelGOSTCommand");
			PushButton pushButton = ribbonPanel.AddItem(pushButtonData) as PushButton;
			pushButton.ToolTip = "Экспортировать модель в спецификацию по ГОСТ в Excel";
			pushButton.LargeImage = Bitmaps.Convert(Properties.Resources.ButtonIcon);
		}
	}

	[Transaction(TransactionMode.Manual)]
	[Regeneration(RegenerationOption.Manual)]
	public class RevitToExcelGOSTCommand : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			Rvt.Windows = new RvtWindows();
			Rvt.Windows.RunLoadingWindow();

			Work.Book = new WorkBook();
			Rvt.Handler = new RvtHandler(commandData, elements);
			Rvt.Control = new RvtControl();
			Rvt.Data = new RvtData();
			Work.Bitmaps = new Bitmaps();

			Rvt.Windows.CloseLoadingWindow();
			Rvt.Windows.RunMainWindow();

			return Rvt.Handler == null ? Result.Failed : Rvt.Handler.Result;
		}
	}
}
