using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Autodesk.Revit.DB;

namespace RevToGOSTv0
{
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;

			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;

			CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(PickedElements.ItemsSource);
			PropertyGroupDescription groupDescription = new PropertyGroupDescription("Category.Name");
			view.GroupDescriptions.Add(groupDescription);
		}

		private void btn_Click(object sender, RoutedEventArgs e)
		{
			Rvt.Data.AvailableElements.UpdateCollection();
			Rvt.Data.PickedElements.UpdateCollection();
		}

		//public void UpdateCategoryListBoxes()
		//{
		//	//AvailableCategories.ItemsSource = Rvt.Data.GetPrintable(Rvt.Data.AvailableCategories);
		//	AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
		//	PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
		//}

		//private void AddCategory_Click(object sender, RoutedEventArgs e)
		//{
		//	////Log.WriteLine("Click! Selected item: {0}", (string)AvailableCategories.SelectedItem);
		//	//Category obj = Rvt.Data.GetCategoryObject(Rvt.Data.AvailableCategories, (string)AvailableCategories.SelectedItem);
		//	//if (obj == null)
		//	//	return;
		//	//////Log.WriteLine("Object: {0}", (string)obj[0]);
		//	//Rvt.Data.RemoveCategoryFromList(Rvt.Data.AvailableCategories, obj);
		//	//Rvt.Data.AddCategoryToList(Rvt.Data.PickedCategories, obj);
		//	//UpdateCategoryListBoxes();
		//}

		//private void Export_Click(object sender, RoutedEventArgs e)
		//{
		//	Rvt.Data.LogPickedElements();
		//}
	}
}
