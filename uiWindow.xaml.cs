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
	/// <summary>
	/// Interaction logic for uiWindow.xaml
	/// </summary>
	public partial class uiWindow : Window
	{
		IList<Element> elements;
		public uiWindow(IList<Element> list)
		{
			InitializeComponent();
			
			elements = list;
			REV2GOST.ItemsSource = elements;
		}

		void NormalNewWindow(object sender, RoutedEventArgs e)
		{
			Window myOwnedWindow = new Window();
			myOwnedWindow.Owner = this;
			myOwnedWindow.Show();
		}
		/*
		private void Window_Click(object sender, RoutedEventArgs e)
		{
			textBox.Text = "LOL";
		}
		*/
		private void Button_Click(object sender, RoutedEventArgs e)
		{
			textBox.Text += "LOL";
		}
	}
}
