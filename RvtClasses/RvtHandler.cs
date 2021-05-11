using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitToGOST
{
	static partial class Rvt
	{
		public static RvtHandler Handler;
	}

	class RvtHandler
	{
		public ExternalCommandData ExtData { get; set; }
		public UIApplication UIApp { get; set; }
		public Document Doc { get; set; }
		public ProjectInfo ProjInfo { get; set; }
		public Result Result { get; set; } = Result.Cancelled;

		public RvtHandler(ExternalCommandData extData, ElementSet elemSet)
		{
			ExtData = extData;
			UIApp = ExtData.Application;
			Doc = UIApp.ActiveUIDocument.Document;
			ProjInfo = Doc.ProjectInformation;
		}
	}
}
