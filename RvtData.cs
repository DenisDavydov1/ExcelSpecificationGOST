using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
	static partial class Rvt
	{
		public static RvtData Data;
	}

	class RvtData
	{
		/*
		** Member properties
		*/

		public class CategoryNode
		{
			public Category Category;
			public ElementCollection Elements;

			public CategoryNode(Category category, ElementCollection elements)
			{
				Category = category;
				Elements = elements;
			}

			public string Name { get { return Category.Name; } }
			public int Id { get { return Category.Id.IntegerValue; } }
			public int Count { get { return Elements.Count; } }
			public override string ToString() { return String.Format("{0} ({1})", Name, Count); }
		}

		public class CategoryNodeCollection : ObservableCollection<CategoryNode>
		{
			protected override void InsertItem(int index, CategoryNode categoryNode)
			{
				base.InsertItem(index, categoryNode);

				if (this == Rvt.Data.PickedCategories)
					Rvt.Data.PickedElements.InsertElementCollection(0, categoryNode.Elements); // index to do
			}

			protected override void RemoveItem(int index)
			{
				Category catToDel = this[index].Category;
				base.RemoveItem(index);
				Rvt.Data.PickedElements.RemoveCategory(catToDel);
				Rvt.Data.AvailableElements.RemoveCategory(catToDel);
			}
		}

		public CategoryNodeCollection AvailableCategories { get; set; }
		public CategoryNodeCollection PickedCategories { get; set; }
		public ElementCollection AvailableElements { get; set; }
		public ElementCollection PickedElements { get; set; }
		public ElementCollection ExportElements { get; set; }

		public int Count { get { return AvailableCategories.Count + PickedCategories.Count; } }

		/*
		** Member methods
		*/

		public RvtData()
		{
			AvailableCategories = new CategoryNodeCollection();
			PickedCategories = new CategoryNodeCollection();
			AvailableElements = new ElementCollection();
			PickedElements = new ElementCollection();
			ExportElements = new ElementCollection();
			InitData();
		}

		public CategoryNodeCollection InitData()
		{
			AvailableCategories = new CategoryNodeCollection();
			foreach (BuiltInCategory enumCat in Enum.GetValues(typeof(BuiltInCategory)))
			{
				try
				{
					FilteredElementCollector collector = new FilteredElementCollector(Rvt.Handler.Doc);
					ElementCollection elemC = new ElementCollection();
					elemC.InsertElementCollection(0, collector.OfCategory(enumCat).ToList());
					Category category = Category.GetCategory(Rvt.Handler.Doc, enumCat);
					if (elemC != null && elemC.Count > 0 &&
						category != null && category.CategoryType == CategoryType.Model)
					{
						AvailableCategories.Add(new CategoryNode(category, elemC));
					}
				}
				catch { continue; }
			}
			return AvailableCategories;
		}

		public CategoryNodeCollection AddCategoryNodeToList(CategoryNodeCollection list, CategoryNode categoryNode)
		{
			if (CategoryNodeInList(list, categoryNode) == true)
				return list;
			list.Add(categoryNode);
			return list;
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
				ExportElements = PickedElements;
			ExportElements.Enumerate();
		}

		public void CreateLines()
		{
			if (ConfFile.FillLine[(int)Work.Book.Table] == null)
				return;
			foreach (ElementContainer elemCont in ExportElements)
			{
				ConfFile.FillLine[(int)Work.Book.Table](elemCont);
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

	} // class RvtData

} // namespace RevitToGOST
