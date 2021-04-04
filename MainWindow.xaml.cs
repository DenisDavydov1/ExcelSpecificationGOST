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
		public MainWindow()
		{
			InitializeComponent();
			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;

			// if (checkbox group in categories) then this code:
			//CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(PickedElements.ItemsSource);
			//PropertyGroupDescription groupDescription = new PropertyGroupDescription("Category.Name");
			//view.GroupDescriptions.Add(groupDescription);
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
	} // class MainWindow

} // namespace RevitToGOST
