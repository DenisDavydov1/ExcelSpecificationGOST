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
		public int Amount { get; set; } = 0;
		public string InstanceName
		{
			get
			{
				return GOST_21_110_2013.ElementName(this);
			}
		}
		public string Type
		{
			get
			{
				return Element.Name;
			}
		}
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

		public bool Equals(ElementContainer other)
		{
			if (this.InstanceName == other.InstanceName &&
				this.Type == other.Type)
				return true;
			return false;
		}
	}

	class ElementCollection : ObservableCollection<ElementContainer>
	{
		/*
		** Work with Element
		*/

		protected void InsertItem(int index, Element element)
		{
			base.InsertItem(index, new ElementContainer(element));
		}

		public void InsertElementCollection(int index, List<Element> elemL)
		{
			foreach (Element elem in elemL)
			{
				this.InsertItem(index++, elem);
			}
		}

		/*
		** Work with ElementContainer
		*/

		protected override void InsertItem(int index, ElementContainer elementContainer)
		{
			if (HasDuplicate(elementContainer) == false)
				base.InsertItem(index, elementContainer);
		}

		public void InsertElementCollection(int index, ElementCollection elemC)
		{
			foreach (ElementContainer elem in elemC)
			{
				if (HasDuplicate(elem) == false)
					this.InsertItem(index++, elem);
			}
		}

		public void AddElementCollection(ElementCollection elemCol)
		{
			foreach (ElementContainer elemCont in elemCol)
			{
				if (HasDuplicate(elemCont) == false)
					this.Add(elemCont);
			}
		}

		/*
		** Common
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

	} // class ElementCollection

} // namespace RevitToGOST
