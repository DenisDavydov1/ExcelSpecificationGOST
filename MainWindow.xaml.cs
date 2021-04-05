using System;
using System.Collections.Generic;
using System.IO;
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
using Microsoft.Win32;

namespace RevitToGOST
{
	public partial class MainWindow : Window
	{
		/*
		** Member properties
		*/

		private CollectionView PickedElementsView { get; set; }
		private CollectionView AvailableElementsView { get; set; }

		/*
		** Member methods
		*/

		public MainWindow()
		{
			InitializeComponent();
			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;
		}

		private void Export_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				Rvt.Control.ExportButton();
			}
			catch (Exception ex)
			{
				Log.WriteLine("Exception caught((: {0}", ex.Message);
			}
		}

		/*
		** Группировать элементы CheckBox
		*/

		private void GroupElemsCheckBox_Checked(object sender, RoutedEventArgs e)
		{
			PickedElementsView = (CollectionView)CollectionViewSource.GetDefaultView(PickedElements.ItemsSource);
			AvailableElementsView = (CollectionView)CollectionViewSource.GetDefaultView(AvailableElements.ItemsSource);
			PickedElementsView.GroupDescriptions.Add(new PropertyGroupDescription("Element.Category.Name"));
			AvailableElementsView.GroupDescriptions.Add(new PropertyGroupDescription("Element.Category.Name"));
			Rvt.Control.GroupElemsCheckBox = true;
		}

		private void GroupElemsCheckBox_Unchecked(object sender, RoutedEventArgs e)
		{
			PickedElementsView.GroupDescriptions.Remove(PickedElementsView.GroupDescriptions.Last());
			AvailableElementsView.GroupDescriptions.Remove(AvailableElementsView.GroupDescriptions.Last());
			Rvt.Control.GroupElemsCheckBox = false;
		}

	} // class MainWindow

} // namespace RevitToGOST
