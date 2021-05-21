using System;
using System.Linq;
using Autodesk.Revit.DB;

namespace ExcelSpecificationGOST
{
	static partial class Rvt
	{
		public static RvtData Data;
	}

	class RvtData
	{
		public CategoryNodeCollection AllCategories { get; set; }
		public CategoryNodeCollection AvailableCategories { get; set; }
		public CategoryNodeCollection PickedCategories { get; set; }
		public ElementCollection AvailableElements { get; set; }
		public ElementCollection PickedElements { get; set; }
		public ElementCollection ExportElements { get; set; }
		public int Count { get { return AvailableCategories.Count + PickedCategories.Count; } }

		public RvtData()
		{
			AllCategories = new CategoryNodeCollection();
			AvailableCategories = new CategoryNodeCollection();
			PickedCategories = new CategoryNodeCollection();
			AvailableElements = new ElementCollection();
			PickedElements = new ElementCollection();
			InitData();
			InitAvailableCategories();
		}

		private void InitData()
		{
			foreach (BuiltInCategory enumCat in Enum.GetValues(typeof(BuiltInCategory)))
			{
				try
				{
					FilteredElementCollector collector = new FilteredElementCollector(Rvt.Handler.Doc);
					ElementCollection elemC = new ElementCollection(collector.OfCategory(enumCat).ToList());
					Category category = Category.GetCategory(Rvt.Handler.Doc, enumCat);
					if (elemC != null && elemC.Count > 0 &&
						category != null && category.CategoryType == CategoryType.Model)
					{
						AllCategories.Add(new CategoryNode(category, elemC));
					}
				}
				catch { continue; }
			}
		}

		private void InitAvailableCategories()
		{
			foreach (CategoryNode node in AllCategories)
				AvailableCategories.Add(node);
		}

		public CategoryNode RemoveCategoryNodeFromList(CategoryNodeCollection list, CategoryNode categoryNode)
		{
			if (CategoryNodeInList(list, categoryNode) == false)
				return null;
			if (list.Remove(categoryNode) == false)
				return null;
			return categoryNode;
		}

		public bool CategoryNodeInList(CategoryNodeCollection list, CategoryNode categoryNode)
		{
			foreach (var node in list)
				if (node.Category.Equals(categoryNode.Category))
					return true;
			return false;
		}

		public void SetExportElements()
		{
			if (Rvt.Control.GroupElemsCheckBox == true)
				ExportElements = PickedElements.GroupByCategory();
			else
				ExportElements = new ElementCollection(PickedElements);
			ExportElements.Enumerate();
		}

		public void FillLines()
		{
			// Fill table
			if (Work.Book.Table != GOST.Standarts.None &&
				ConfFile.FillLine[(int)Work.Book.Table] != null)
			{
				foreach (ElementContainer elemCont in ExportElements)
				{
					ConfFile.FillLine[(int)Work.Book.Table](elemCont);
				}
				Work.Book.ConvertElementCollectionsToLists();
			}
		}

		public void InitExportElements()
		{
			ExportElements = new ElementCollection();
		}

		public void InsertColumnsEnumerationLines()
		{
			if (Rvt.Control.EnumerateColumnsCheckBox == false)
				return;
			for (int i = 0; i < ExportElements.Count; i += ConfFile.Lines[(int)Work.Book.Table])
			{
				ExportElements.Insert(i, new ElementContainer(ElementContainer.ContType.ColumnsEnumeration));
			}
		}
	}
}
