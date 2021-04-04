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
	class GOST_21_110_2013 : IGostData
	{
		/*
		**	Member properties
		*/

		public class Line
		{
			// "Поз."
			public int Position;

			// "Наименование и техническая характеристика"
			public string Name;

			// "Тип, марка, обозначение документа, опросного листа"
			public string Type;

			// "Код продукции"
			public string ProdCode; //int ProdCode;

			// "Поставщик"
			public string Provider;

			// "Ед. измерения"
			public string Unit;

			// "Количество"
			public double Amount;

			// "Масса 1 ед., кг"
			public double Weight;

			// "Примечание"
			public string Note;

			//public Line(int pos) { }
		}

		//public ElementSet ElemSet { get; set; }
		public ElementCollection ElemCol { get; set; }
		public List<Line> Lines { get; set; }
		private int Position;


		/*
		**	Member methods
		*/

		public GOST_21_110_2013(ElementCollection elemCol)
		{
			ElemCol = elemCol;
			Position = 1;
		}

		public void FillLines()
		{
			Lines = new List<Line>();

			foreach (Element elem in ElemCol)
			{
				if (ElementHasDuplicate(elem, ElemCol) == true)
					continue;
				Line line = new Line();

				///// "Поз."
				line.Position = Position++;

				///// "Наименование и техническая характеристика"
				//// Variable: public string Name;
				/// Real parameters:
				// ELEM_FAMILY_PARAM		// Family
				// SYMBOL_FAMILY_NAME_PARAM	// Family Name
				// ALL_MODEL_FAMILY_NAME	// Family Name

				/// Doubtful:
				// FAMILY_NAME_PSEUDO_PARAM		// Family
				// DPART_ORIGINAL_FAMILY		// Original Family

				line.Name = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.ELEM_FAMILY_PARAM,
					BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM,
					BuiltInParameter.ALL_MODEL_FAMILY_NAME
				});


				///// "Тип, марка, обозначение документа, опросного листа"
				//// Variable: public string Type;
				/// Real parameters:
				// ELEM_TYPE_PARAM						// Type
				// SYMBOL_NAME_PARAM					// Type Name
				// ALL_MODEL_TYPE_NAME					// Type Name
				// ELEM_FAMILY_AND_TYPE_PARAM			// Family and Type
				// SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM	// Family and Type

				line.Type = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.ELEM_TYPE_PARAM,
					BuiltInParameter.SYMBOL_NAME_PARAM,
					BuiltInParameter.ALL_MODEL_TYPE_NAME,
					BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
					BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM
				});

				///// "Код продукции"
				//// Variable: public string ProdCode; //int ProdCode;
				/// Real parameters:
				// STRUCTURAL_FAMILY_CODE_NAME	// Code Name
				// OMNICLASS_CODE				// OmniClass Number
				// FABRICATION_PRODUCT_CODE		// Product Code
				// UNIFORMAT_CODE				// Assembly code

				line.ProdCode = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.STRUCTURAL_FAMILY_CODE_NAME,
					BuiltInParameter.OMNICLASS_CODE,
					BuiltInParameter.FABRICATION_PRODUCT_CODE,
					BuiltInParameter.UNIFORMAT_CODE
				});

				///// "Поставщик"
				//// Variable: public string Provider;
				/// Real parameters:
				// ALL_MODEL_MANUFACTURER	// Manufacturer
				// FABRICATION_VENDOR		// Vendor
				// FABRICATION_VENDOR_CODE	// Vendor Code

				line.Provider = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.ALL_MODEL_MANUFACTURER,
					BuiltInParameter.FABRICATION_VENDOR,
					BuiltInParameter.FABRICATION_VENDOR_CODE
				});

				///// "Ед. измерения"
				//// Variable: public string Unit;
				/// Real parameters:
				// IMPORT_DISPLAY_UNITS				// Import Units
				// ALTERNATE_UNITS					// Alternate Units
				// POINT_ELEMENT_MEASUREMENT_TYPE	// Measurement Type

				line.Unit = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.IMPORT_DISPLAY_UNITS,
					BuiltInParameter.ALTERNATE_UNITS,
					BuiltInParameter.POINT_ELEMENT_MEASUREMENT_TYPE
				});

				///// "Количество"
				//// Variable: public double Amount;
				/// Real parameters:

				line.Amount = 1.0;

				///// "Масса 1 ед., кг"
				//// Variable: public double Weight;
				/// Real parameters:
				// FABRIC_SHEET_MASSUNIT	// "Sheet Mass per Unit Area": Structural Sheet Mass
				// per Unit Area [Sheet Mass / (Overall Length * Overall Width)]
				// COUPLER_WEIGHT			// Mass

				line.Weight = GetDoubleParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.FABRIC_SHEET_MASSUNIT,
					BuiltInParameter.COUPLER_WEIGHT
				});

				///// "Примечание"
				//// Variable: public string Note;
				/// Real parameters:
				// ALL_MODEL_INSTANCE_COMMENTS	// Comments
				// ALL_MODEL_DESCRIPTION		// Description
				// ALL_MODEL_TYPE_COMMENTS		// Type comments
				// MARKUPS_NOTES				// Notes
				// SHEET_ASSEMBLY_KEYNOTE		// Assembly: Keynote
				// FABRICATION_PART_NOTES		// Fabrication Notes
				// ALL_MODEL_URL				// URL

				line.Note = GetStringParameter(elem, new BuiltInParameter[] {
					BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
					BuiltInParameter.ALL_MODEL_DESCRIPTION,
					BuiltInParameter.ALL_MODEL_TYPE_COMMENTS,
					BuiltInParameter.MARKUPS_NOTES,
					BuiltInParameter.SHEET_ASSEMBLY_KEYNOTE,
					BuiltInParameter.FABRICATION_PART_NOTES,
					BuiltInParameter.ALL_MODEL_URL
				});

				Lines.Add(line);
			}
		}

		// TO DO ! INDEX ERROR
		private bool ElementHasDuplicate(Element newElem, ElementCollection elemCol)
		{
			//for (int i = 0; i < elemCol.Count && elemCol[i] != newElem; ++i)
			//{
			//	if (elemCol[i].Equals(newElem))
			//	{
			//		Lines[i].Amount += 1.0;
			//		return true;
			//	}
			//}
			return false;
		}

		//private bool EqualElements(Element elem1, Element elem2)
		//{
		//	if (elem1.get_Parameter)
		//	return false;
		//}

		private static string GetStringParameter(Element elem, BuiltInParameter[] parameters)
		{
			foreach (BuiltInParameter param in parameters)
			{
				string paramValue = GetStringParameterValue(elem.get_Parameter(param));
				if (paramValue != null && paramValue.Length > 0)
					return paramValue;
			}
			return String.Empty;
		}

		private static double GetDoubleParameter(Element elem, BuiltInParameter[] parameters)
		{
			foreach (BuiltInParameter param in parameters)
			{
				double paramValue = GetDoubleParameterValue(elem.get_Parameter(param));
				if (paramValue > 0.0)
					return paramValue;
			}
			return 0.0;
		}

		//private static string GetStringParameterValue(Parameter param)
		//{
		//	if (param == null || param.HasValue == false)
		//		return String.Empty;

		//	// Storage type: None
		//	if (param.StorageType == StorageType.None)
		//		return String.Empty;

		//	// Storage type: Integer
		//	if (param.StorageType == StorageType.Integer)
		//		return param.AsInteger().ToString();

		//	// Storage type: Double
		//	if (param.StorageType == StorageType.Double)
		//		return param.AsDouble().ToString();

		//	// Storage type: String
		//	if (param.StorageType == StorageType.String)
		//	{
		//		string asString = param.AsString();
		//		string asValueString = param.AsValueString();
		//		if (asValueString != null && asValueString.Length > 0)
		//			return asValueString;
		//		if (asString != null && asString.Length > 0)
		//			return asString;
		//		return String.Empty;
		//	}

		//	// Storage type: ElementId
		//	if (param.StorageType == StorageType.ElementId)
		//		return param.AsElementId().IntegerValue.ToString();

		//	//return param.AsValueString();
		//	return String.Empty;
		//}

		private static string GetStringParameterValue(Parameter param)
		{
			if (param == null || param.HasValue == false)
				return String.Empty;

			string output = param.AsValueString();
			if (output != null && output.Length > 0)
				return output;

			output = param.AsString();
			if (output != null && output.Length > 0)
				return output;

			return String.Empty;
		}

		private static double GetDoubleParameterValue(Parameter param)
		{
			if (param == null || param.HasValue == false)
				return 0.0;

			double output = param.AsDouble();
			if (output > 0.000001)
				return output;

			output = param.AsInteger();
			if (output > 0.000001)
				return output;

			return 0.0;
		}

		public List<List<string>> FillList()
		{
			List<List<string>> output = new List<List<string>>();
			//for (int i = 0; i < Lines.Count; ++i)
			foreach (Line line in Lines)
			{
				output.Add(new List<string>());
				List<string> curLineList = output.Last();

				curLineList.Add(line.Position.ToString());  // "Поз."
				curLineList.Add(line.Name);                 // "Наименование и техническая характеристика"
				curLineList.Add(line.Type);                 // "Тип, марка, обозначение документа, опросного листа"
				curLineList.Add(line.ProdCode);             // "Код продукции"
				curLineList.Add(line.Provider);             // "Поставщик"
				curLineList.Add(line.Unit);                 // "Ед. измерения"
				curLineList.Add(line.Amount.ToString());    // "Количество"
				curLineList.Add(line.Weight.ToString());    // "Масса 1 ед., кг"
				curLineList.Add(line.Note);                 // "Примечание"
			}
			return output;
		}
	} // class GOST_21_110_2013

} // namespace RevitToGOST
