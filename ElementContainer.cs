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
	public class ElementContainer
	{
		/*
		** Member properties
		*/

		public Element Element { get; set; }
		public string InstanceName { get { return GOST_21_110_2013.ElementName(this); } }
		public string Type { get { return Element.Name; } }
		public int Amount { get; set; } = 0;
		public Category Category { get { return Element.Category; } }
		public string CategoryName { get { return Category.Name; } }
		public string CategoryLine { get; set; } = null;
		public int Position { get; set; } = 0;
		public ContType LineType { get; set; } = ContType.Element;
		public List<string> Line { get; set; }

		public enum ContType
		{
			Element,
			Category,
			ColumnsEnumeration
		}

		/*
		** Member methods
		*/

		public ElementContainer(ContType contType = ContType.Element)
		{
			Amount = 1;
			LineType = contType;
		}
		
		public ElementContainer(Element element)
		{
			Element = element;
			Amount = 1;
			LineType = ContType.Element;
		}

		public ElementContainer(string categoryLine)
		{
			CategoryLine = categoryLine;
			LineType = ContType.Category;
		}

		public ElementContainer(ElementContainer other)
		{
			Element = other.Element;
			Amount = other.Amount;
			CategoryLine = other.CategoryLine;
			Position = other.Position;
			LineType = other.LineType;
			Line = other.Line;
		}

		public bool Equals(ElementContainer other)
		{
			if (this.InstanceName == other.InstanceName &&
				this.Type == other.Type)
				return true;
			return false;
		}

	} // class ElementContainer

	public class ElementCollection : ObservableCollection<ElementContainer>
	{
		/*
		** Member properties
		*/

		public int ElementCount
		{
			get
			{
				int count = 0;
				foreach (ElementContainer elemCont in this)
					count += elemCont.Amount;
				return count;
			}
		}

		/*
		** Constructors
		*/

		public ElementCollection() { }

		public ElementCollection(List<Element> elemList)
		{
			foreach (Element elem in elemList)
			{
				ElementContainer elemCont = new ElementContainer(elem);
				if (HasDuplicate(elemCont) == false)
					base.Add(elemCont);
			}
		}

		public ElementCollection(ElementCollection other)
		{
			foreach (ElementContainer elem in other)
				base.Add(elem);
				//base.Add(new ElementContainer(elem));
		}

		/*
		** Insert ElementContainer(s)
		*/

		public void InsertElementCollection(int index, ElementCollection elemC)
		{
			foreach (ElementContainer elem in elemC)
			{
				base.InsertItem(index++, elem);
			}
		}

		public void AddElementCollection(ElementCollection elemCol)
		{
			foreach (ElementContainer elemCont in elemCol)
			{
				this.Add(elemCont);
			}
		}

		/*
		** Member methods
		*/

		protected override void RemoveItem(int index)
		{
			base.RemoveItem(index);
		}

		public void RemoveCategory(Category catToDel)
		{
			for (int i = 0; i < Count; i++)
			{
				if (this[i].Element.Category.Id.IntegerValue == catToDel.Id.IntegerValue)
					RemoveAt(i--);
			}
		}

		private bool HasDuplicate(ElementContainer elementContainer)
		{
			if (elementContainer.LineType != ElementContainer.ContType.Element)
				return false;
			foreach (ElementContainer currentElem in this)
			{
				if (currentElem.LineType != ElementContainer.ContType.Element)
					continue;
				if (currentElem.Equals(elementContainer))
				{
					currentElem.Amount++;
					return true;
				}
			}
			return false;
		}

		public ElementCollection GroupByCategory()
		{
			// Fill a dictionary by elements categories
			Dictionary<string, ElementCollection> categoryDict = new Dictionary<string, ElementCollection>();
			foreach (ElementContainer elemCont in Rvt.Data.PickedElements)
			{
				if (categoryDict.Keys.Contains(elemCont.CategoryName) == false)
					categoryDict.Add(elemCont.CategoryName, new ElementCollection());
				categoryDict[elemCont.CategoryName].Add(elemCont);
			}

			// Create a new element collection filled by the dictionary order
			ElementCollection groupedCol = new ElementCollection();
			foreach (var pair in categoryDict)
			{
				groupedCol.Add(new ElementContainer(pair.Key));	// Add category line
				groupedCol.AddElementCollection(pair.Value);	// Add elements of category
			}
			return groupedCol;
		}

		public void Enumerate()
		{
			int pos = 1;
			foreach (ElementContainer elemCont in this)
			{
				if (elemCont.LineType == ElementContainer.ContType.Element)
					elemCont.Position = pos++;
				else if (elemCont.LineType == ElementContainer.ContType.Category)
					pos = 1;
			}
		}

		/*
		** Sort items
		*/

		public enum SortBy
		{
			InstanceName,
			Type,
			Amount
		}

		public void Sort(SortBy byWhat)
		{
			if (IsAscendingOrdered(byWhat) == false)
				Quicksort(this, 0, Count - 1, true, byWhat);
			else if (IsDescendingOrdered(byWhat) == false)
				Quicksort(this, 0, Count - 1, false, byWhat);
		}

		private void Quicksort(ElementCollection array, int start, int end, bool ascending, SortBy byWhat)
		{
			if (start >= end)
				return;
			int pivot = Partition(array, start, end, ascending, byWhat);
			Quicksort(array, start, pivot - 1, ascending, byWhat);
			Quicksort(array, pivot + 1, end, ascending, byWhat);
		}

		private int Partition(ElementCollection array, int start, int end, bool ascending, SortBy byWhat)
		{
			int marker = start; //divides left and right subarrays
			for (int i = start; i < end; i++)
			{
				//array[end] is pivot
				if (byWhat == SortBy.InstanceName)
				{
					if ((ascending == true && array[i].InstanceName.CompareTo(array[end].InstanceName) < 0) ||
						(ascending == false && array[i].InstanceName.CompareTo(array[end].InstanceName) > 0))
					{
						Swap(marker, i);
						marker += 1;
					}
				}
				else if (byWhat == SortBy.Type)
				{
					if ((ascending == true && array[i].Type.CompareTo(array[end].Type) < 0) ||
						(ascending == false && array[i].Type.CompareTo(array[end].Type) > 0))
					{
						Swap(marker, i);
						marker += 1;
					}
				}
				else // if sort by elements amount
				{
					if ((ascending == true && array[i].Amount < array[end].Amount) ||
						(ascending == false && array[i].Amount > array[end].Amount))
					{
						Swap(marker, i);
						marker += 1;
					}
				}
			}
			// put pivot(array[end]) between left and right subarrays
			Swap(marker, end);
			return marker;
		}

		private void Swap(int index1, int index2)
		{
			ElementContainer tmp = this[index1];
			this[index1] = this[index2];
			this[index2] = tmp;
		}

		private bool IsAscendingOrdered(SortBy byWhat)
		{
			if (byWhat == SortBy.InstanceName)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].InstanceName.CompareTo(this[i + 1].InstanceName) > 0)
						return false;
			}
			else if (byWhat == SortBy.Type)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Type.CompareTo(this[i + 1].Type) > 0)
						return false;
			}
			else
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Amount > this[i + 1].Amount)
						return false;
			}
			return true;
		}

		private bool IsDescendingOrdered(SortBy byWhat)
		{
			if (byWhat == SortBy.InstanceName)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].InstanceName.CompareTo(this[i + 1].InstanceName) < 0)
						return false;
			}
			else if (byWhat == SortBy.Type)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Type.CompareTo(this[i + 1].Type) < 0)
						return false;
			}
			else
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Amount < this[i + 1].Amount)
						return false;
			}
			return true;
		}

	} // class ElementCollection

} // namespace RevitToGOST
