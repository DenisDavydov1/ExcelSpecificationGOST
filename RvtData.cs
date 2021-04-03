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

		public class CategoryNodeCollection : ObservableCollection<CategoryNode>
		{
			//public new void Add(CategoryNode categoryNode)
			//{
			//	Log.WriteLine("New Add!!!");
			//	base.Add(categoryNode);
			//}

			//public new void Insert(int index, CategoryNode categoryNode)
			//{
			//	base.Insert(index, categoryNode);
			//	Log.WriteLine("New Insert!!!");
			//}

			public CategoryNode GetNodeByCategory(Category category)
			{
				foreach (CategoryNode node in this)
					if (node.Category.Equals(category))
						return node;
				//for (int i = 0; i < Count; ++i)
				//	if (this[i].Category.Equals(category))
				//		return this[i];
				return null;
			}

			protected override void InsertItem(int index, CategoryNode categoryNode)
			{
				base.InsertItem(index, categoryNode);
				//Log.WriteLine("Insert {0} to {1}", categoryNode.Name, index);

				if (this == Rvt.Data.PickedCategories && categoryNode.PickedElements.Count == 0)
					GostTools.AddToObservableCollection(categoryNode.AvailableElements, categoryNode.PickedElements);
			}

			protected override void RemoveItem(int index)
			{
				base.RemoveItem(index);
				//Log.WriteLine("Remove from {0}", index);
			}
		}

		public class ElementCollection : ObservableCollection<Element>
		{
			protected override void InsertItem(int index, Element element)
			{
				base.InsertItem(index, element);

				//if (this == Rvt.Data.AvailableElements)
				//{
				//	CategoryNode node = Rvt.Data.PickedCategories.GetNodeByCategory(element.Category);
				//	if (node == null) // this element's category not picked
				//		return;
				//	node.AvailableElements.Add(element);
				//}
			}

			public void UpdateCollection()
			{
				//ElementCollection elements;
				//if (this == Rvt.Data.AvailableElements)
				//	elements = Rvt.Data.AvailableElements;
				this.Clear();
				if (this == Rvt.Data.AvailableElements)
				{
					foreach (CategoryNode node in Rvt.Data.PickedCategories)
						foreach (Element element in node.AvailableElements)
							this.Add(element);
				}
				else if (this == Rvt.Data.PickedElements)
				{
					foreach (CategoryNode node in Rvt.Data.PickedCategories)
						foreach (Element element in node.PickedElements)
							this.Add(element);
				}
			}
		}

		public class CategoryNode
		{
			public Category Category;
			public ObservableCollection<Element> AvailableElements;
			public ObservableCollection<Element> PickedElements;

			public CategoryNode(Category category, ObservableCollection<Element> elements)
			{
				Category = category;
				AvailableElements = elements;
				PickedElements = new ObservableCollection<Element>();
			}

			public string Name { get { return Category.Name; } }
			public int Id { get { return Category.Id.IntegerValue; } }
			public int Count { get { return AvailableElements.Count + PickedElements.Count; } }
			public override string ToString() { return String.Format("{0} ({1})", Name, Count); }
		}

		public CategoryNodeCollection AvailableCategories { get; set; }
		public CategoryNodeCollection PickedCategories { get; set; }
		public ElementCollection AvailableElements { get; set; }
		public ElementCollection PickedElements { get; set; }

		//public ObservableCollection<Element> AvailableElements { get; set; }
		//public ObservableCollection<Element> PickedElements { get; set; }

		/*
		** Member methods
		*/

		public RvtData()
		{
			AvailableCategories = new CategoryNodeCollection();
			PickedCategories = new CategoryNodeCollection();
			AvailableElements = new ElementCollection();
			PickedElements = new ElementCollection();
			InitData();
			//Log.WriteLine("Avail cat: {0}", AvailableCategories.Count);
		}

		public CategoryNodeCollection InitData()
		{
			AvailableCategories = new CategoryNodeCollection();
			foreach (BuiltInCategory enumCat in Enum.GetValues(typeof(BuiltInCategory)))
			{
				try
				{
					FilteredElementCollector collector = new FilteredElementCollector(Rvt.Handler.Doc);
					ObservableCollection<Element> elemList = new ObservableCollection<Element>(collector.OfCategory(enumCat));
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

		private int CompareCategoryNodes(CategoryNode x, CategoryNode y)
		{
			string name1 = x.Name;
			string name2 = y.Name;
			return String.Compare(name1, name2);
		}

		//public bool CategoryNodeInList(List<CategoryNode> list, BuiltInCategory category)
		//{
		//	foreach (var node in list)
		//		if (node.Id() == (int)category)
		//			return true;
		//	return false;
		//}

		public bool CategoryNodeInList(CategoryNodeCollection list, CategoryNode categoryNode)
		{
			foreach (var node in list)
				if (node.Category.Equals(categoryNode.Category))
					return true;
			return false;
		}

		public int Count()
		{
			return AvailableCategories.Count() + PickedCategories.Count();
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

		public void LogPickedElements()
		{
			Log.WriteLine("\n\nPicked elems:\n");
			foreach (var catNode in PickedCategories)
				Log.WriteLine(String.Join("", catNode.PickedElements));
		}

	} // class RvtData
} // namespace RevToGOSTv0
