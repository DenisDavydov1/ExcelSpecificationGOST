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
		public static void FillLine(ElementContainer elemCont)
		{
			//// Заголовок категории
			if (elemCont.LineType == ElementContainer.ContType.Category)
			{
				elemCont.Line = new List<string>() { "", elemCont.CategoryLine, "", "", "", "", "", "", "" };
			}

			//// Нумерация стоблцов
			else if (elemCont.LineType == ElementContainer.ContType.ColumnsEnumeration)
			{
				elemCont.Line = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9" };
			}

			//// Настоящий элемент
			else
			{
				elemCont.Line = new List<string>();

				// "Поз."
				elemCont.Line.Add(elemCont.Position.ToString());

				// "Наименование и техническая характеристика"
				elemCont.Line.Add(ElementName(elemCont));

				// "Тип, марка, обозначение документа, опросного листа"
				elemCont.Line.Add(ElementType(elemCont));

				// "Код продукции"
				elemCont.Line.Add(ElementProdCode(elemCont));

				// "Поставщик"
				elemCont.Line.Add(ElementProvider(elemCont));

				// "Ед. измерения"
				string unit = ElementUnit(elemCont);
				if (unit == String.Empty)
					elemCont.Line.Add("Шт.");
				else
					elemCont.Line.Add(unit);

				// "Количество"
				double amount = ElementAmount(elemCont);
				if (amount == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(amount.ToString());

				// "Масса 1 ед., кг"
				double weight = ElementWeight(elemCont);
				if (weight == 0.0)
					elemCont.Line.Add(String.Empty);
				else
					elemCont.Line.Add(weight.ToString());

				// "Примечание"
				elemCont.Line.Add(ElementNote(elemCont));
			}
		}

		public static string ElementName(ElementContainer elemContainer)
		{
			///// "Наименование и техническая характеристика"
			//// Variable: public string Name;
			/// Real parameters:
			// ELEM_FAMILY_PARAM		// Family
			// SYMBOL_FAMILY_NAME_PARAM	// Family Name
			// ALL_MODEL_FAMILY_NAME	// Family Name

			/// Doubtful:
			// FAMILY_NAME_PSEUDO_PARAM		// Family
			// DPART_ORIGINAL_FAMILY		// Original Family

			return GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.ELEM_FAMILY_PARAM,
				BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM,
				BuiltInParameter.ALL_MODEL_FAMILY_NAME
			});
		}

		public static string ElementType(ElementContainer elemContainer)
		{
			///// "Тип, марка, обозначение документа, опросного листа"
			//// Variable: public string Type;
			/// Real parameters:
			// ELEM_TYPE_PARAM						// Type
			// SYMBOL_NAME_PARAM					// Type Name
			// ALL_MODEL_TYPE_NAME					// Type Name
			// ELEM_FAMILY_AND_TYPE_PARAM			// Family and Type
			// SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM	// Family and Type

			string type = GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.ELEM_TYPE_PARAM,
				BuiltInParameter.SYMBOL_NAME_PARAM,
				BuiltInParameter.ALL_MODEL_TYPE_NAME,
				BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
				BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM
			});
			if (type == String.Empty)
				type = elemContainer.Element.Category.Name;
			return type;
		}

		public static string ElementProdCode(ElementContainer elemContainer)
		{
			///// "Код продукции"
			//// Variable: public string ProdCode; //int ProdCode;
			/// Real parameters:
			// STRUCTURAL_FAMILY_CODE_NAME	// Code Name
			// OMNICLASS_CODE				// OmniClass Number
			// FABRICATION_PRODUCT_CODE		// Product Code
			// UNIFORMAT_CODE				// Assembly code

			return GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.STRUCTURAL_FAMILY_CODE_NAME,
				BuiltInParameter.OMNICLASS_CODE,
				BuiltInParameter.FABRICATION_PRODUCT_CODE,
				BuiltInParameter.UNIFORMAT_CODE
			});
		}

		public static string ElementProvider(ElementContainer elemContainer)
		{
			///// "Поставщик"
			//// Variable: public string Provider;
			/// Real parameters:
			// ALL_MODEL_MANUFACTURER	// Manufacturer
			// FABRICATION_VENDOR		// Vendor
			// FABRICATION_VENDOR_CODE	// Vendor Code

			return GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.ALL_MODEL_MANUFACTURER,
				BuiltInParameter.FABRICATION_VENDOR,
				BuiltInParameter.FABRICATION_VENDOR_CODE
			});
		}

		public static string ElementUnit(ElementContainer elemContainer)
		{
			///// "Ед. измерения"
			//// Variable: public string Unit;
			/// Real parameters:
			// IMPORT_DISPLAY_UNITS				// Import Units
			// ALTERNATE_UNITS					// Alternate Units
			// POINT_ELEMENT_MEASUREMENT_TYPE	// Measurement Type

			return GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.IMPORT_DISPLAY_UNITS,
				BuiltInParameter.ALTERNATE_UNITS,
				BuiltInParameter.POINT_ELEMENT_MEASUREMENT_TYPE
			});
		}

		public static double ElementAmount(ElementContainer elemContainer)
		{
			///// "Количество"
			//// Variable: public double Amount;
			/// Real parameters:

			return elemContainer.Amount;
			
			// TO DO! Consider measurement units!

		}

		public static double ElementWeight(ElementContainer elemContainer)
		{
			///// "Масса 1 ед., кг"
			//// Variable: public double Weight;
			/// Real parameters:
			// FABRIC_SHEET_MASSUNIT	// "Sheet Mass per Unit Area": Structural Sheet Mass
			// per Unit Area [Sheet Mass / (Overall Length * Overall Width)]
			// COUPLER_WEIGHT			// Mass

			return GetDoubleParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.FABRIC_SHEET_MASSUNIT,
				BuiltInParameter.COUPLER_WEIGHT
			});
		}

		public static string ElementNote(ElementContainer elemContainer)
		{
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

			return GetStringParameter(elemContainer.Element, new BuiltInParameter[] {
				BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
				BuiltInParameter.ALL_MODEL_DESCRIPTION,
				BuiltInParameter.ALL_MODEL_TYPE_COMMENTS,
				BuiltInParameter.MARKUPS_NOTES,
				BuiltInParameter.SHEET_ASSEMBLY_KEYNOTE,
				BuiltInParameter.FABRICATION_PART_NOTES,
				BuiltInParameter.ALL_MODEL_URL
			});
		}

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

	} // class GOST_21_110_2013

} // namespace RevitToGOST
