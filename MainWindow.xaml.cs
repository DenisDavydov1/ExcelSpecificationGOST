using System;
using System.Collections.Generic;
using System.ComponentModel;
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
			
			Rvt.Windows.Condition = RvtWindows.Status.Idle;
			Rvt.Windows.ConditionChanged += new PropertyChangedEventHandler(ConditionChangeHandler);

			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;
			MakeAllComboBoxUpdate();
			DrawPreview();
		}

		private void DrawPreview()
		{
			//ImageTable.Source = PreviewImages.Images[(int)Work.Book.Table];
			ImageTable.Source = new BitmapImage(new Uri(@"pack://application:,,,/RevitToGOST;component/Previews/GOST_21_110_2013_Table1.png"));
		}

		private void ConditionChangeHandler(object sender, PropertyChangedEventArgs e)
		{
			if (Rvt.Windows.Condition == RvtWindows.Status.Idle)
			{
				ExportProgressBar.Value = 0.0;
				EnableControls(true);
				AvailableCategories.Items.Refresh();
				PickedCategories.Items.Refresh();
				AvailableElements.Items.Refresh();
				PickedElements.Items.Refresh();
			}
			else if (Rvt.Windows.Condition == RvtWindows.Status.Export)
				EnableControls(false);
			else if (Rvt.Windows.Condition == RvtWindows.Status.Error)
			{
				MessageBox.Show(String.Format("{0}\n\nStack trace:\n{1}",
					Rvt.Control.LastException.Message, Rvt.Control.LastException.StackTrace),
						"Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}
			else if (Rvt.Windows.Condition == RvtWindows.Status.Sort)
			{
				EnableControls(false);
			}
		}

		private void EnableControls(bool state = true)
		{
			// No tab
			Export.IsEnabled = state;
			GroupElemsCheckBox.IsEnabled = state;

			// Tab 1
			TitleComboBox.IsEnabled = state;
			TableComboBox.IsEnabled = state;
			StampComboBox.IsEnabled = state;
			DopComboBox.IsEnabled = state;
			EnumerateColumnsCheckBox.IsEnabled = state;

			// Tab 2
			AvailableCategories.IsEnabled = state;
			PickedCategories.IsEnabled = state;

			// Tab 3
			AvailableElements.IsEnabled = state;
			PickedElements.IsEnabled = state;
		}


		/////// Export button and worker methods ///////

		private void Export_Click(object sender, RoutedEventArgs e)
		{
			Rvt.Windows.Condition = RvtWindows.Status.Export;

			BackgroundWorker exportWorker = new BackgroundWorker();
			exportWorker.WorkerReportsProgress = true;
			exportWorker.DoWork += ExportWorker_DoWork;
			exportWorker.ProgressChanged += ExportWorker_ProgressChanged;
			exportWorker.RunWorkerCompleted += ExportWorker_RunWorkerCompleted;
			exportWorker.RunWorkerAsync();
		}

		private void ExportWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			try
			{
				Rvt.Control.ExportProcedure(sender);
			}
			catch (Exception ex)
			{
				Rvt.Control.LastException = ex;
				Rvt.Windows.Condition = RvtWindows.Status.Error;
			}
		}

		private void ExportWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			ExportProgressBar.Value = e.ProgressPercentage;
		}

		private void ExportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Error != null || Rvt.Windows.Condition == RvtWindows.Status.Error)
			{
				Rvt.Windows.Condition = RvtWindows.Status.Error;
			}
			else
			{
				Rvt.Control.SaveProcedure();    // Save file
				Rvt.Data.InitExportElements();	// Unset Rvt.Data.ExportElements
			}
			Rvt.Windows.Condition = RvtWindows.Status.Idle;
		}

		/////// /////// /////// /////// /////// ///////


		private void huy_Click(object sender, RoutedEventArgs e)
		{

		}


		/////// Группировать элементы (CheckBox) ///////
		///
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
		
		/////// /////// /////// /////// /////// ///////
		

		////////////////////////////////////// TAB 1 //////////////////////////////////////

		/*
		** Настройка конфигурации таблицы (ComboBox)
		*/

		private void MakeAllComboBoxUpdate()
		{
			TitleComboBox_SelectionChanged(null, null);
			TableComboBox_SelectionChanged(null, null);
			StampComboBox_SelectionChanged(null, null);
			DopComboBox_SelectionChanged(null, null);
		}

		private void TitleComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (TitleComboBox.SelectedIndex == 0)
				Work.Book.Title = GOST.Standarts.None;
			else if (TitleComboBox.SelectedIndex == 1)
				Work.Book.Title = GOST.Standarts.GOST_P_21_101_2020_Title_12;
			// DrawPreview(); TO DO!
		}

		private void TableComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (TableComboBox.SelectedIndex == 0)
				Work.Book.Table = GOST.Standarts.GOST_21_110_2013_Table1;
			DrawPreview();
		}

		private void StampComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (StampComboBox.SelectedIndex == 0)
				Work.Book.Stamp = GOST.Standarts.None;
			else if (StampComboBox.SelectedIndex == 1)
				Work.Book.Stamp = GOST.Standarts.GOST_P_21_101_2020_Stamp3;
			// DrawPreview(); TO DO!
		}

		private void DopComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (DopComboBox.SelectedIndex == 0)
				Work.Book.Dop = GOST.Standarts.None;
			else if (DopComboBox.SelectedIndex == 1)
				Work.Book.Dop = GOST.Standarts.GOST_P_21_101_2020_Dop3;
			// DrawPreview(); TO DO!
		}

		////////////////////////////////////// TAB 2 //////////////////////////////////////

		private void PickAllCategoriesButton_Click(object sender, RoutedEventArgs e)
		{
			while (Rvt.Data.AvailableCategories.Count > 0)
			{
				CategoryNode tmp = Rvt.Data.AvailableCategories[0];
				Rvt.Data.AvailableCategories.RemoveAt(0);
				Rvt.Data.PickedCategories.Insert(Rvt.Data.PickedCategories.Count, tmp);
			}
		}

		private void ReleaseAllCategoriesButton_Click(object sender, RoutedEventArgs e)
		{
			while (Rvt.Data.PickedCategories.Count > 0)
			{
				CategoryNode tmp = Rvt.Data.PickedCategories[0];
				Rvt.Data.PickedCategories.RemoveAt(0);
				Rvt.Data.AvailableCategories.Insert(Rvt.Data.AvailableCategories.Count, tmp);
			}
		}

		private void PickAllElementsButton_Click(object sender, RoutedEventArgs e)
		{
			while (Rvt.Data.AvailableElements.Count > 0)
			{
				ElementContainer tmp = Rvt.Data.AvailableElements[0];
				Rvt.Data.AvailableElements.RemoveAt(0);
				Rvt.Data.PickedElements.Add(tmp);
			}
		}

		private void ReleaseAllElementsButton_Click(object sender, RoutedEventArgs e)
		{
			while (Rvt.Data.PickedElements.Count > 0)
			{
				ElementContainer tmp = Rvt.Data.PickedElements[0];
				Rvt.Data.PickedElements.RemoveAt(0);
				Rvt.Data.AvailableElements.Add(tmp);
			}
		}

		////////////////////////////////////// TAB 3 //////////////////////////////////////
		////////////////////////////////////// TAB 4 //////////////////////////////////////

		/*
		** Нумерация столбцов (CheckBox)
		*/

		private void EnumerateColumnsCheckBox_Checked(object sender, RoutedEventArgs e)
		{
			Rvt.Control.EnumerateColumnsCheckBox = true;
			// DrawPreview(); TO DO!
		}

		private void EnumerateColumnsCheckBox_Unchecked(object sender, RoutedEventArgs e)
		{
			Rvt.Control.EnumerateColumnsCheckBox = false;
			// DrawPreview(); TO DO!
		}

		/*
		** All tabs gridview sort function
		*/

		private void CommonColumnHeader_Click(object sender, RoutedEventArgs e)
		{
			Rvt.Windows.Condition = RvtWindows.Status.Sort;
			try
			{
				string name = ((GridViewColumnHeader)sender).Name;
				// Tab 1 available
				if (name == "AvailableCategoriesNameHeader")
					Rvt.Data.AvailableCategories.Sort(true);
				else if (name == "AvailableCategoriesCountHeader")
					Rvt.Data.AvailableCategories.Sort(false);

				// Tab 1 picked
				else if (name == "PickedCategoriesNameHeader")
					Rvt.Data.PickedCategories.Sort(true);
				else if (name == "PickedCategoriesCountHeader")
					Rvt.Data.PickedCategories.Sort(false);

				// Tab 2 available
				else if (name == "AvailableElementsInstanceNameHeader")
					Rvt.Data.AvailableElements.Sort(ElementCollection.SortBy.InstanceName);
				else if (name == "AvailableElementsTypeHeader")
					Rvt.Data.AvailableElements.Sort(ElementCollection.SortBy.Type);
				else if (name == "AvailableElementsAmountHeader")
					Rvt.Data.AvailableElements.Sort(ElementCollection.SortBy.Amount);

				// Tab 2 picked
				else if (name == "PickedElementsInstanceNameHeader")
					Rvt.Data.PickedElements.Sort(ElementCollection.SortBy.InstanceName);
				else if (name == "PickedElementsTypeHeader")
					Rvt.Data.PickedElements.Sort(ElementCollection.SortBy.Type);
				else if (name == "PickedElementsAmountHeader")
					Rvt.Data.PickedElements.Sort(ElementCollection.SortBy.Amount);
				Rvt.Windows.Condition = RvtWindows.Status.Idle;
			}
			catch (Exception ex)
			{
				Rvt.Control.LastException = ex;
				Rvt.Windows.Condition = RvtWindows.Status.Error;
			}
			Rvt.Windows.Condition = RvtWindows.Status.Idle;
		}

	} // class MainWindow

} // namespace RevitToGOST
