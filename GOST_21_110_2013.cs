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

		public GOST_21_110_2013(ref WorkBook workbook, string worksheetName = Constants.DefaultName)
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
			this.Worksheet.PageSetup.Margins.Top = 0.5 / Constants.inch;
			this.Worksheet.PageSetup.Margins.Left = 2 / Constants.inch;
			this.Worksheet.PageSetup.Margins.Right = 0.5 / Constants.inch;
			this.Worksheet.PageSetup.Margins.Bottom = 0.5 / Constants.inch;
			this.Worksheet.PageSetup.Margins.Footer = 0;
			this.Worksheet.PageSetup.Margins.Header = 0;
			this.Worksheet.PageSetup.CenterHorizontally = false;
			this.Worksheet.PageSetup.CenterVertically = false;
		}

		public override void FormatHeader()
		{
			// Set height
			this.Worksheet.Row(1).Height = 32 * Constants.mm_h;
			for (int i = 2; i < 20; i++)
				this.Worksheet.Row(i).Height = 8 * Constants.mm_h;

			// Set width
			this.Worksheet.Column(1).Width = 20 * Constants.mm_w;
			this.Worksheet.Column(2).Width = 130 * Constants.mm_w;
			this.Worksheet.Column(3).Width = 60 * Constants.mm_w;
			this.Worksheet.Column(4).Width = 35 * Constants.mm_w;
			this.Worksheet.Column(5).Width = 45 * Constants.mm_w;
			this.Worksheet.Column(6).Width = 20 * Constants.mm_w;
			this.Worksheet.Column(7).Width = 20 * Constants.mm_w;
			this.Worksheet.Column(8).Width = 25 * Constants.mm_w;
			this.Worksheet.Column(9).Width = 40 * Constants.mm_w;

			// Set font
			this.Worksheet.Style.Font.FontName = "GOST A";
			//this.Worksheet.Column(1).Style.Font.FontName = "GOST A";
			this.Worksheet.Column(1).Style.Font.FontSize = 7;
			this.Worksheet.Column(1).Style.Font.Italic = true;

			// Set border
			this.Worksheet.Range("A1:I41").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
			this.Worksheet.Range("A1:I41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("A1:I1").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("A1:A41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("B1:B41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("C1:C41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("D1:D41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("E1:E41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("F1:F41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("G1:G41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("H1:H41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
			this.Worksheet.Range("I1:I41").Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

			// Fill header
			XMLTools.FillCells(this.Worksheet, "Наименование фурнитуры", XLAlignmentHorizontalValues.Center, XLAlignmentVerticalValues.Center,
				"GOST A", "bold italic", 16, "A1:A1", XLBorderStyleValues.Medium);
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
