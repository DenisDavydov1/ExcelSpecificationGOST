using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;

namespace RevToGOSTv0
{
	class AbstractGOST : GOST
	{
		public AbstractGOST(WorkBook workbook, string worksheetName = Constants.DefaultName)
		{
			this.WorksheetName = worksheetName;
			this.Workbook = workbook;
			this.Worksheet = this.Workbook.AddWorkSheet(this.WorksheetName);
			this.SetFormat();

			// Set font
			this.Worksheet.Style.Font.FontName = "GOST A";
			//this.Worksheet.Column(1).Style.Font.FontName = "GOST A";
			this.Worksheet.Column(1).Style.Font.FontSize = 7;
			this.Worksheet.Column(1).Style.Font.Italic = true;

			this.SetSizes();
			this.ReshapeTable();

			//this.SetFormat();
			//this.FormatHeader();

			//this.Lines = new List<List<string>>();
		}
		public override void FillLines(List<Element> elements)
		{
			throw new NotImplementedException();
		}

		public override void FillTable()
		{
			throw new NotImplementedException();
		}

		public override void FormatHeader()
		{
			throw new NotImplementedException();
		}

		public override void SetFormat()
		{
			// Page setup
			this.Worksheet.PageSetup.PageOrientation = XLPageOrientation.Landscape;
			this.Worksheet.PageSetup.AdjustTo(100);
			this.Worksheet.PageSetup.PaperSize = XLPaperSize.A3Paper;
			this.Worksheet.PageSetup.VerticalDpi = 600;
			this.Worksheet.PageSetup.HorizontalDpi = 600;

			// Margins setup
			this.Worksheet.PageSetup.Margins.Top = 0;
			this.Worksheet.PageSetup.Margins.Left = 0;
			this.Worksheet.PageSetup.Margins.Right = 0;
			this.Worksheet.PageSetup.Margins.Bottom = 0;
			this.Worksheet.PageSetup.Margins.Footer = 0;
			this.Worksheet.PageSetup.Margins.Header = 0;
			this.Worksheet.PageSetup.CenterHorizontally = false;
			this.Worksheet.PageSetup.CenterVertically = false;
		}

		private void SetSizes()
		{
			//this.ColumnsSize = new List<int[]>() {
			//	new int[] { 7 },
			//	new int[] { 2, 1, 3, 1 },
			//	new int[] { 1, 1, 1, 3, 1 },
			//	new int[] { 7 }
			//};
			//this.RowsSize = new List<int[]>() {
			//	new int[] { 5 },
			//	new int[] { 2, 2, 1 },
			//	new int[] { 1, 3, 1 },
			//	new int[] { 1, 3, 1 },
			//	new int[] { 5 }
			//};

			// 8 x 25
			// 137
			//this.ColumnsSize = new List<int[]>() {
			//	new int[] { 420 },														// 5
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 37 // header
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 45 // line 1
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 53 // line 2
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 61 // line 3
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 69 // line 4
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 77 // line 5
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 85 // line 6
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 93 // line 7
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 101 // line 8
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 109 // line 9
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 117 // line 10
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 125 // line 11
			//	new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },				// 133 // line 12
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 141 // line 13 + dop grafa
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 149 // line 14
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 152
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 157 // line 15
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 165 // line 16
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 167
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 173 // line 17
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 181 // line 18
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 187
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 189 // line 19
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 197 // line 20
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 205 // line 21
			//	new int[] { 5, 5, 5, 5, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },	// 207 // end of dop grafa
			//	//new int[] { 8, 5, 7, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },		// dop grafa 2
			//	new int[] { 8, 5, 7, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },		// 213 // line 22
			//	new int[] { 8, 5, 7, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },		// 221 // line 23
			//	new int[] { 8, 5, 7, 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 },		// 229 // line 24
			//	new int[] { 8, 5, 7, 5 },												// 232 // dop gr
			//	new int[] { 5, 410, 5 },												// 237 // empty line 25
			//	new int[] { 8, 5, 7, 210, 10, 10, 10, 10, 15, 10, 120, 5 },				// 242 // st
			//	new int[] { 8, 5, 7, 210, 10, 10, 10, 10, 15, 10, 120, 5 },				// 247 // st
			//	new int[] { 8, 5, 7, 210, 10, 10, 10, 10, 15, 10, 120, 5 },				// 252 // st
			//	new int[] { 8, 5, 7, 210, 10, 10, 10, 10, 15, 10, 120, 5 },				// 257 // st
			//	new int[] { 8, 5, 7, 210, 10, 10, 10, 10, 15, 10, 120, 5 },				// 262 // st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 15, 15, 20, 5 },			// 267 // dop gr + st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 15, 15, 20, 5 },			// 272 // st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 15, 15, 20, 5 },			// 277 // st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 50, 5 },					// 282 // st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 50, 5 },					// 287 // st
			//	new int[] { 8, 5, 7, 210, 20, 20, 15, 10, 70, 50, 5 },					// 292 // st dop last
			//	new int[] { 420 }														// 297
			//};
			//this.RowsSize = new List<int[]>() {
			//	new int[] { 297 },												// 5 // field
			//	new int[] { 142, 65, 90 },			// 8 // field
			//	new int[] { 142, 65, 25, 35, 25, 5 },			// 10 // field
			//	new int[] { 142, 10, 15, 20, 20, 25, 35, 25, 5 },			// 13 // field
			//	new int[] { 142, 10, 15, 20, 20, 25, 35, 25, 5 },			// 15 // field
			//	new int[] { 142, 10, 15, 20, 20, 25, 35, 25, 5 },			// 20 // field
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 63, 5 },	// 40 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 63, 5 },			// 170 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 63, 5 },			// 230 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 240 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 250 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 260 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 265 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 270 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 285 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 },			// 295 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 15, 15, 5 },			// 310 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 15, 15, 5 },			// 330 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 15, 15, 5 },			// 350 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 15, 15, 5 },			// 365 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 5, 10, 5, 10, 5 },			// 375 // lines
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 5, 10, 5, 10, 5 },			// 380 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 5, 10, 5, 10, 5 },			// 395 // st
			//	new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 10, 15, 5, 10, 5, 10, 5 },			// 415 // st & last
			//	new int[] { 297 }
			//};

			this.ColumnsSize = new List<int[]>() {
				new int[] { 420 },
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // header
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 1
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 2
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 3
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 4
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 5
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 6
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 7
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 8
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 9
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 10
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 11
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 12
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 13
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 14
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 15
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 16
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 17
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 18
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 19
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 20
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 21
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 22
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 23
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 24
				new int[] { 20, 20, 130, 60, 35, 45, 20, 20, 25, 40, 5 }, // line 25
				new int[] { 20, 395, 5 },
				new int[] { 420 }
			};
			this.RowsSize = new List<int[]>() {
				new int[] { 297 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 5, 32, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 55, 5 },
				new int[] { 297 }
			};

			//this.ColumnsSize = new List<int[]>() {
			//	new int[] { 150 },
			//	new int[] { 20, 13, 50, 35, 10, 17, 5 }, // header
			//	new int[] { 20, 13, 50, 35, 10, 17, 5 }, // header
			//	new int[] { 20, 13, 50, 35, 10, 17, 5 }, // header
			//	new int[] { 20, 13, 50, 35, 10, 17, 5 }, // header
			//	new int[] { 20, 125, 5 },
			//	new int[] { 150 }
			//};
			//this.RowsSize = new List<int[]>() {
			//	new int[] { 100 },
			//	new int[] { 5, 10, 5, 5, 5, 65, 5 },
			//	new int[] { 5, 10, 5, 5, 5, 65, 5 },
			//	new int[] { 5, 10, 5, 5, 5, 65, 5 },
			//	new int[] { 5, 10, 5, 5, 5, 65, 5 },
			//	new int[] { 100 }
			//};



		}

		private void ReshapeTable()
		{
			SortedSet<int> columns = new SortedSet<int>();
			SortedSet<int> rows = new SortedSet<int>();

			for (int i = 0; i < this.ColumnsSize.Count; i++)
				for (int j = 0; j < this.ColumnsSize[i].Length; j++)
					columns.Add(this.ColumnsSize[i].Take(j + 1).Sum());
			columns.Remove(0);
			Log.WriteLine(String.Join(", ", columns));

			for (int i = 0; i < this.RowsSize.Count; i++)
				for (int j = 0; j < this.RowsSize[i].Length; j++)
					rows.Add(this.RowsSize[i].Take(j + 1).Sum());
			rows.Remove(0);
			Log.WriteLine(String.Join(", ", rows));

			for (int i = 0, size = columns.ElementAt(0); i < columns.Count; i++)
			{
				if (i > 0)
					size = columns.ElementAt(i) - columns.ElementAt(i - 1);
				this.Worksheet.Column(i + 1).Width = XMLTools.mmToWidth(size);
			}
			for (int i = 0, size = rows.ElementAt(0); i < rows.Count; i++)
			{
				if (i > 0)
					size = rows.ElementAt(i) - rows.ElementAt(i - 1);
				this.Worksheet.Row(i + 1).Height = XMLTools.mmToHeight(size);
			}

			Array columns_arr = columns.ToArray();
			for (int i = 0; i < this.ColumnsSize.Count; i++)
			{
				for (int j = 0; j < this.ColumnsSize[i].Length; j++)
				{
					int j_real = Array.IndexOf(columns_arr, this.ColumnsSize[i].Take(j + 1).Sum()) + 1;
					if (j_real == 0 || this.ColumnsSize[i].Take(j + 1).Sum() >= columns.Max())
						break;
					this.Worksheet.Cell(i + 1, j_real).Style.Border.RightBorder = XLBorderStyleValues.Thin;
				}
			}

			Array rows_arr = rows.ToArray();
			for (int i = 0; i < this.RowsSize.Count; i++)
			{
				for (int j = 0; j < this.RowsSize[i].Length; j++)
				{
					int j_real = Array.IndexOf(rows_arr, this.RowsSize[i].Take(j + 1).Sum()) + 1;
					//Log.WriteLine("j: {0}, j_real: {1}, i: {2}, rows_max: {3}", j, j_real, i, rows.Max());
					if (j_real == 0 || this.RowsSize[i].Take(j + 1).Sum() >= rows.Max())
						break;
					this.Worksheet.Cell(j_real, i + 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
				}
			}
		}
	}
}
