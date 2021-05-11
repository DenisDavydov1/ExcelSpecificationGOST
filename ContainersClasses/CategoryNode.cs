using System;
using System.Collections.ObjectModel;
using Autodesk.Revit.DB;

namespace RevitToGOST
{
	public class CategoryNode
	{
		public Category Category { get; set; }
		public ElementCollection Elements { get; set; }
		public string Name { get { return Category.Name; } }
		public int Id { get { return Category.Id.IntegerValue; } }
		public int Count { get { return Elements.ElementCount; } }

		public CategoryNode(Category category, ElementCollection elements)
		{
			Category = category;
			Elements = elements;
		}

		public override string ToString() { return String.Format("{0} ({1})", Name, Count); }
	}

	public class CategoryNodeCollection : ObservableCollection<CategoryNode>
	{
		public CategoryNodeCollection() { }
		public CategoryNodeCollection(CategoryNodeCollection other)
		{
			foreach (CategoryNode node in other)
				Add(node);
		}
		protected override void InsertItem(int index, CategoryNode categoryNode)
		{
			base.InsertItem(index, categoryNode);
			if (this == Rvt.Data.PickedCategories)
				Rvt.Data.PickedElements.InsertElementCollection(Rvt.Data.PickedElements.Count, this[index].Elements);
		}

		public new void Insert(int index, CategoryNode categoryNode)
		{
			InsertItem(index, categoryNode);
		}

		public new void Add(CategoryNode categoryNode)
		{
			base.InsertItem(Count, categoryNode);
		}

		protected override void RemoveItem(int index)
		{
			Category catToDel = this[index].Category;
			base.RemoveItem(index);
			Rvt.Data.PickedElements.RemoveCategory(catToDel);
			Rvt.Data.AvailableElements.RemoveCategory(catToDel);
		}

		#region sort items

		public void Sort(bool byName)
		{
			if (IsAscendingOrdered(byName) == false)
				Quicksort(this, 0, Count - 1, true, byName);
			else if (IsDescendingOrdered(byName) == false)
				Quicksort(this, 0, Count - 1, false, byName);
		}

		private void Quicksort(CategoryNodeCollection array, int start, int end, bool ascending, bool byName)
		{
			if (start >= end)
				return;
			int pivot = Partition(array, start, end, ascending, byName);
			Quicksort(array, start, pivot - 1, ascending, byName);
			Quicksort(array, pivot + 1, end, ascending, byName);
		}

		private int Partition(CategoryNodeCollection array, int start, int end, bool ascending, bool byName)
		{
			int marker = start; //divides left and right subarrays
			for (int i = start; i < end; i++)
			{
				//array[end] is pivot
				if (byName == true)
				{
					if ((ascending == true && array[i].Name.CompareTo(array[end].Name) < 0) ||
						(ascending == false && array[i].Name.CompareTo(array[end].Name) > 0))
					{
						Swap(marker, i);
						marker += 1;
					}
				}
				else // if sort by elements count
				{
					if ((ascending == true && array[i].Count < array[end].Count) ||
						(ascending == false && array[i].Count > array[end].Count))
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
			CategoryNode tmp = this[index1];
			this[index1] = this[index2];
			this[index2] = tmp;
		}

		private bool IsAscendingOrdered(bool byName)
		{
			if (byName == true)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Name.CompareTo(this[i + 1].Name) > 0)
						return false;
			}
			else
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Count > this[i + 1].Count)
						return false;
			}
			return true;
		}

		private bool IsDescendingOrdered(bool byName)
		{
			if (byName == true)
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Name.CompareTo(this[i + 1].Name) < 0)
						return false;
			}
			else
			{
				for (int i = 0; i < Count - 1; ++i)
					if (this[i].Count < this[i + 1].Count)
						return false;
			}
			return true;
		}

		#endregion
	}
}
