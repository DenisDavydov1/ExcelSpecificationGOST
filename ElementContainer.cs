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
		public int Amount { get; set; }
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

		/*
		** Member methods
		*/

		public ElementContainer()
		{
			Amount = 1;
		}
		public ElementContainer(Element element)
		{
			Element = element;
			Amount = 1;
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
			foreach (ElementContainer currentElem in this)
			{
				if (currentElem.Equals(elementContainer))
				{
					currentElem.Amount++;
					return true;
				}
			}
			return false;
		}

		

	} // class ElementCollection
}
