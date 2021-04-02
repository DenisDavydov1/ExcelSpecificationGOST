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

namespace RevToGOSTv0
{
	static partial class Rvt
	{
		public static RvtData Data;
	}

	class RvtData
	{
		/*
		** Member fields
		*/

		public class CategoryNode
		{
			public Category Category;
			public List<Element> AvailableElements;
			public List<Element> PickedElements;

			//public CategoryNode() {}
			public CategoryNode(Category category, List<Element> elements)
			{
				Category = category;
				//AvailableElements = new List<Element>(elements);
				AvailableElements = elements;
				PickedElements = new List<Element>();
			}

			public string Name() { return Category.Name; }
			public int Id() { return Category.Id.IntegerValue; }
			public int Count() { return AvailableElements.Count() + PickedElements.Count(); }

		}

		public List<CategoryNode> AvailableCategories { get; set; }
		public List<CategoryNode> PickedCategories { get; set; }

		/*
		** Member methods
		*/

		public RvtData()
		{
			AvailableCategories = new List<CategoryNode>();
			PickedCategories = new List<CategoryNode>();
			InitData();
		}

		public List<CategoryNode> InitData()
		{
			AvailableCategories = new List<CategoryNode>();
			foreach (BuiltInCategory enumCat in Enum.GetValues(typeof(BuiltInCategory)))
			{
				try
				{
					FilteredElementCollector collector = new FilteredElementCollector(Rvt.Handler.Doc);
					List<Element> elemList = collector.OfCategory(enumCat).ToList();
					Category category = Category.GetCategory(Rvt.Handler.Doc, enumCat);
					if (elemList != null && elemList.Count > 0 &&
						category != null && category.CategoryType == CategoryType.Model)
					{
						AvailableCategories.Add(new CategoryNode(category, elemList));
					}
				}
				catch { continue; }
			}
			return AvailableCategories;
		}

		public List<CategoryNode> AddCategoryNodeToList(List<CategoryNode> list, CategoryNode categoryNode)
		{
			if (CategoryNodeInList(list, categoryNode) == true)
				return list;
			list.Add(categoryNode);
			return list;
		}

		public CategoryNode RemoveCategoryNodeFromList(List<CategoryNode> list, CategoryNode categoryNode)
		{
			if (CategoryNodeInList(list, categoryNode) == false)
				return null;
			if (list.Remove(categoryNode) == false)
				return null;
			return categoryNode;
		}

		private int CompareCategoryNodes(CategoryNode x, CategoryNode y)
		{
			string name1 = x.Name();
			string name2 = y.Name();
			return String.Compare(name1, name2);
		}

		//public bool CategoryNodeInList(List<CategoryNode> list, BuiltInCategory category)
		//{
		//	foreach (var node in list)
		//		if (node.Id() == (int)category)
		//			return true;
		//	return false;
		//}

		public bool CategoryNodeInList(List<CategoryNode> list, CategoryNode categoryNode)
		{
			foreach (var node in list)
				if (node.Category.Equals(categoryNode.Category))
					return true;
			return false;
		}

		//public List<string> GetPrintable(List<Category> list)
		//{
		//	List<string> output = new List<string>();
		//	foreach (var category in list)
		//		output.Add(String.Format("{0} ({1})", category.Name, GetCategoryElements(category).Count));
		//	return output;
		//}

		//public List<Element> GetCategoryElements(Category category)
		//{
		//	FilteredElementCollector collector = new FilteredElementCollector(Rvt.Handler.Doc);
		//	return collector.OfCategory((BuiltInCategory)category.Id.IntegerValue).ToList();
		//}

		//public Category GetCategoryObject(List<Category> list, string name)
		//{
		//	if (name == null || name.Length == 0)
		//		return null;
		//	name = name.Substring(0, name.LastIndexOf(" ("));
		//	foreach (var category in list)
		//		if (category.Name == name)
		//			return category;
		//	return null;
		//}

		//private void PlaceCategoryInRightOrder(object[] category, List<>)

	} // class RvtData
} // namespace RevToGOSTv0
