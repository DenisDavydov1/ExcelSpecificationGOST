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

	class ElementCollection : ObservableCollection<Element>
	{
		protected override void InsertItem(int index, Element element)
		{
			base.InsertItem(index, element);
		}

		protected override void RemoveItem(int index)
		{
			base.RemoveItem(index);
		}

		public void InsertElementCollection(int index, ElementCollection elemC)
		{
			foreach (Element elem in elemC)
			{
				this.Insert(index++, elem);
			}
		}

		public void InsertElementCollection(int index, List<Element> elemL)
		{
			foreach (Element elem in elemL)
			{
				this.Insert(index++, elem);
			}
		}

		internal void RemoveCategory(Category catToDel)
		{
			for (int i = 0; i < Count; i++)
			{
				if (this[i].Category.Id.IntegerValue == catToDel.Id.IntegerValue)
					RemoveAt(i--);
			}
		}
	}

	class RvtData
	{
		/*
		** Member fields
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

		private int CompareCategoryNodes(CategoryNode x, CategoryNode y)
		{
			string name1 = x.Name;
			string name2 = y.Name;
			return String.Compare(name1, name2);
		}

		public bool CategoryNodeInList(CategoryNodeCollection list, CategoryNode categoryNode)
		{
			foreach (var node in list)
				if (node.Category.Equals(categoryNode.Category))
					return true;
			return false;
		}

		public int Count { get { return AvailableCategories.Count + PickedCategories.Count; } }

	} // class RvtData
} // namespace RevToGOSTv0








/*
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

public class CategoryNodeCollection : ObservableCollection<CategoryNode>
		{
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

				if (this == Rvt.Data.PickedCategories && categoryNode.PickedElements.Count == 0)
					GostTools.AddToObservableCollection(categoryNode.AvailableElements, categoryNode.PickedElements);
				Rvt.Data.UpdateElementCollections();
			}

			protected override void RemoveItem(int index)
			{
				base.RemoveItem(index);
				Rvt.Data.UpdateElementCollections();
			}

			public void RemoveElement(Element elem)
			{
				foreach (var node in this)
				{
					if (node.AvailableElements.Remove(elem) == true ||
						node.PickedElements.Remove(elem) == true)
						return;
				}
			}

			public void InsertElementToAvailable(int index, Element element)
			{
				int g = 0;
				foreach (CategoryNode node in this)
				{
					for (int i = 0; i <= node.AvailableElements.Count; ++i, ++g)
					{
						if (g == index)
						{
							node.AvailableElements.Insert(i, element);
							return;
						}
					}
				}
			}

			public void InsertElementToPicked(int index, Element element)
			{
				int g = 0;
				foreach (CategoryNode node in this)
				{
					for (int i = 0; i <= node.PickedElements.Count; ++i, ++g)
					{
						if (g == index)
						{
							Log.WriteLine("Insert now: g: {0}, i: {1}", g, i);
							node.PickedElements.Insert(i, element);
							Log.WriteLine("Insert not failed");
							return;
						}
					}
				}
			}
		}

		public class ElementCollection : ObservableCollection<Element>
		{
			//private int GetInsertIndex(int collectionIndex)
			//{
			//	if (this == Rvt.Data.AvailableElements)
			//	{
			//		//int max = Rvt.Data.PickedCategories[0].Count;
			//		foreach (CategoryNode node in Rvt.Data.PickedCategories)
			//		{
			//			for (int index = 0; index < node.AvailableElements.Count; ++index, --collectionIndex)
			//			{
			//				if (collectionIndex == 0)
			//					return index;
			//			}
			//		}
			//	}
			//	return 0;
			//}

			private int GetCategoryStartIndex(Category category)
			{
				int index = 0;

				if (this == Rvt.Data.AvailableElements)
				{
					foreach (var node in Rvt.Data.PickedCategories)
					{
						if (node.Category.Equals(category) == true)
							return index;
						index += node.AvailableElements.Count;
					}
				}
				else if (this == Rvt.Data.PickedElements)
				{
					foreach (var node in Rvt.Data.PickedCategories)
					{
						if (node.Category.Equals(category) == true)
							return index;
						index += node.PickedElements.Count;
					}
				}
				return index;
			}

			protected override void InsertItem(int index, Element element)
			{
				this.Clear();

				////base.InsertItem(index, element);
				if (this == Rvt.Data.PickedElements)
				{
					Log.WriteLine("Picked");
					//Rvt.Data.PickedCategories.InsertElementToPicked(index, element);
					//Log.WriteLine("huicked");
				}
				else if (this == Rvt.Data.AvailableElements)
				{
					Log.WriteLine("Available");
				//	//Rvt.Data.PickedCategories.InsertElementToAvailable(index, element);
				}
			}

			protected override void RemoveItem(int index)
			{
				base.RemoveItem(index);
				Rvt.Data.PickedCategories.RemoveElement(this[index]);
				Rvt.Data.UpdateElementCollections();
			}

			public void UpdateCollection()
			{
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

		public void UpdateElementCollections()
		{
			try
			{
				AvailableElements.UpdateCollection();
				PickedElements.UpdateCollection();
			}
			catch { Log.WriteLine("Exception on update"); }
		}

		public void Logall()
		{
			Log.ClearLog();
			Log.WriteLine("Available Elems");
			foreach (var node in Rvt.Data.PickedCategories)
				foreach (var elem in node.AvailableElements)
					Log.WriteLine("{0}", elem.Name);
			Log.WriteLine("\nPicked Elems");
			foreach (var node in Rvt.Data.PickedCategories)
				foreach (var elem in node.PickedElements)
					Log.WriteLine("{0}", elem.Name);
			//Log.WriteLine("Picked Cats");
			//foreach (var node in Rvt.Data.AvailableCategories)
			//	Log.WriteLine(node.Name);
			//Log.WriteLine("\nPicked Cats");
			//foreach (var node in Rvt.Data.PickedCategories)
			//	Log.WriteLine(node.Name);
		}

	} // class RvtData
} // namespace RevToGOSTv0

*/