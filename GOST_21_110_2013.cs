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
	class GOST_21_110_2013 : GOST
	{

	/*
	**	Member properties
	*/

	/*
	**	Member methods
	*/

		public GOST_21_110_2013(WorkBook workbook, string worksheetName = Constants.DefaultName)
		{
			this.WorksheetName = worksheetName;
			this.Workbook = workbook;
			this.Worksheet = this.Workbook.AddWorkSheet(this.WorksheetName);

			this.SetFormat();
			this.FormatHeader();

			this.Lines = new List< List<string> >();
		}

		public override void SetFormat() // A3_landscape
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

		public override void FormatHeader()
		{
			// Set columns widths
			//this.Worksheet.Column(1).Width = XMLTools.mmToSize(5);  // A
			//this.Worksheet.Column(2).Width = XMLTools.mmToSize(3);  // B
			//this.Worksheet.Column(3).Width = XMLTools.mmToSize(2);      // C
			//this.Worksheet.Column(4).Width = XMLTools.mmToSize(3);    // D
			//this.Worksheet.Column(5).Width = XMLTools.mmToSize(2);      // E
			//this.Worksheet.Column(6).Width = XMLTools.mmToSize(5);  // F
			//this.Worksheet.Column(7).Width = XMLTools.mmToSize(20);     // G
			//this.Worksheet.Column(8).Width = XMLTools.mmToSize(130);        // H
			//this.Worksheet.Column(9).Width = XMLTools.mmToSize(60);     // I
			//this.Worksheet.Column(10).Width = XMLTools.mmToSize(10);      // J
			//this.Worksheet.Column(11).Width = XMLTools.mmToSize(10);      // K
			//this.Worksheet.Column(12).Width = XMLTools.mmToSize(10);      // L
			//this.Worksheet.Column(13).Width = XMLTools.mmToSize(5);      // M
			//this.Worksheet.Column(14).Width = XMLTools.mmToSize(5);      // N
			//this.Worksheet.Column(15).Width = XMLTools.mmToSize(15);      // O
			//this.Worksheet.Column(16).Width = XMLTools.mmToSize(10);      // P
			//this.Worksheet.Column(17).Width = XMLTools.mmToSize(15);      // Q
			//this.Worksheet.Column(18).Width = XMLTools.mmToSize(20);      // R
			//this.Worksheet.Column(19).Width = XMLTools.mmToSize(20);      // S
			//this.Worksheet.Column(20).Width = XMLTools.mmToSize(15);      // T
			//this.Worksheet.Column(21).Width = XMLTools.mmToSize(10);      // U
			//this.Worksheet.Column(22).Width = XMLTools.mmToSize(5);      // V
			//this.Worksheet.Column(23).Width = XMLTools.mmToSize(15);      // W
			//this.Worksheet.Column(24).Width = XMLTools.mmToSize(20);      // X
			//this.Worksheet.Column(25).Width = XMLTools.mmToSize(5);      // Y

			this.Worksheet.Row(1).Height = XMLTools.mmToHeight(1);
			this.Worksheet.Row(2).Height = XMLTools.mmToHeight(2);
			this.Worksheet.Row(3).Height = XMLTools.mmToHeight(3);
			this.Worksheet.Row(4).Height = XMLTools.mmToHeight(4);
			this.Worksheet.Row(5).Height = XMLTools.mmToHeight(5);
			this.Worksheet.Row(6).Height = XMLTools.mmToHeight(6);
			this.Worksheet.Row(7).Height = XMLTools.mmToHeight(7);
			this.Worksheet.Row(8).Height = XMLTools.mmToHeight(8);
			this.Worksheet.Row(9).Height = XMLTools.mmToHeight(9);
			this.Worksheet.Row(10).Height = XMLTools.mmToHeight(10);
			this.Worksheet.Row(11).Height = XMLTools.mmToHeight(11);
			this.Worksheet.Row(12).Height = XMLTools.mmToHeight(12);
			this.Worksheet.Row(13).Height = XMLTools.mmToHeight(13);
			this.Worksheet.Row(14).Height = XMLTools.mmToHeight(14);
			this.Worksheet.Row(15).Height = XMLTools.mmToHeight(15);
			this.Worksheet.Row(16).Height = XMLTools.mmToHeight(20); // 20
			this.Worksheet.Row(17).Height = XMLTools.mmToHeight(25); // 25
			this.Worksheet.Row(18).Height = XMLTools.mmToHeight(30); // 30
			this.Worksheet.Row(19).Height = XMLTools.mmToHeight(35); // 35
			this.Worksheet.Row(20).Height = XMLTools.mmToHeight(40); // 40
			this.Worksheet.Row(21).Height = XMLTools.mmToHeight(45); // 45
			this.Worksheet.Row(22).Height = XMLTools.mmToHeight(50); // 50
			this.Worksheet.Row(23).Height = XMLTools.mmToHeight(55); // 55
			this.Worksheet.Row(24).Height = XMLTools.mmToHeight(60); // 60
			this.Worksheet.Row(25).Height = XMLTools.mmToHeight(70); // 70
			this.Worksheet.Row(26).Height = XMLTools.mmToHeight(80); // 80
			this.Worksheet.Row(27).Height = XMLTools.mmToHeight(90); // 90

			//this.Worksheet.Row(1).Height = 3.088;
			//this.Worksheet.Row(2).Height = 6;
			//this.Worksheet.Row(3).Height = 8.97;
			//this.Worksheet.Row(4).Height = 11.25;
			//this.Worksheet.Row(5).Height = 14.231;
			//this.Worksheet.Row(6).Height = 17.91;
			//this.Worksheet.Row(7).Height = 20.371;
			//this.Worksheet.Row(8).Height = 23.27;
			//this.Worksheet.Row(9).Height = 25.59;
			//this.Worksheet.Row(10).Height = 28.49;
			//this.Worksheet.Row(11).Height = 32.13;
			//this.Worksheet.Row(12).Height = 33.77;
			//this.Worksheet.Row(13).Height = 36.8;
			//this.Worksheet.Row(14).Height = 39.84;
			//this.Worksheet.Row(15).Height = 41.99;
			//this.Worksheet.Row(16).Height = 57.2; // 20
			//this.Worksheet.Row(17).Height = 70.75; // 25
			//this.Worksheet.Row(18).Height = 85; // 30
			//this.Worksheet.Row(19).Height = 99.24; // 35
			//this.Worksheet.Row(20).Height = 113.24; // 40
			//this.Worksheet.Row(21).Height = 127.5; // 45
			//this.Worksheet.Row(22).Height = 141.738; // 50
			//this.Worksheet.Row(23).Height = 155.97; // 55
			//this.Worksheet.Row(24).Height = 170.2; // 60
			//this.Worksheet.Row(25).Height = 197.98; // 70
			//this.Worksheet.Row(26).Height = 226.74; // 80
			//this.Worksheet.Row(27).Height = 255; // 90


			//// Set height
			//this.Worksheet.Row(1).Height = 32 * Constants.mm_h;
			//for (int i = 2; i < 20; i++)
			//	this.Worksheet.Row(i).Height = 8 * Constants.mm_h;

			//// Set width
			//this.Worksheet.Column(1).Width = 20 * Constants.mm_w;
			//this.Worksheet.Column(2).Width = 130 * Constants.mm_w;
			//this.Worksheet.Column(3).Width = 60 * Constants.mm_w;
			//this.Worksheet.Column(4).Width = 35 * Constants.mm_w;
			//this.Worksheet.Column(5).Width = 45 * Constants.mm_w;
			//this.Worksheet.Column(6).Width = 20 * Constants.mm_w;
			//this.Worksheet.Column(7).Width = 20 * Constants.mm_w;
			//this.Worksheet.Column(8).Width = 25 * Constants.mm_w;
			//this.Worksheet.Column(9).Width = 40 * Constants.mm_w;

			//// Set font
			//this.Worksheet.Style.Font.FontName = "GOST A";
			////this.Worksheet.Column(1).Style.Font.FontName = "GOST A";
			//this.Worksheet.Column(1).Style.Font.FontSize = 7;
			//this.Worksheet.Column(1).Style.Font.Italic = true;

			//// Set border
			//this.Worksheet.Range("A1:I41").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
			//this.Worksheet.Range("A1:I41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("A1:I1").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("A1:A41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("B1:B41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("C1:C41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("D1:D41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("E1:E41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("F1:F41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("G1:G41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("H1:H41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			//this.Worksheet.Range("I1:I41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

			//// Fill header
			//XMLTools.FillCells(this.Worksheet, "Наименование фурнитуры", XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center,
			//	"GOST A", "bold italic", 16, "A1:A1", XLBorderStyleValues.Medium);
		}

		public override void FillLines(List<Element> elements)
		{
			//List<string> line = new List<string>();

			foreach (Element elem in elements)
				this.Lines.Add(new List<string>() { elem.Name });
		}

		public override void FillTable()
		{
			int i = 2, j;

			foreach (List<string> line in this.Lines)
			{
				j = 1;
				foreach (string item in line)
				{
					this.Worksheet.Cell(i, j).Value = item;
					j++;
				}
				i++;
			}
		}

	} // class GOST_21_110_2013

} // namespace RevToGOSTv0
