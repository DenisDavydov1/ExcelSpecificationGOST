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

			Rvt.Control.Condition = RvtControl.Status.Idle;
			Rvt.Control.ConditionChanged += new PropertyChangedEventHandler(ConditionChangeHandler);

			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;

			MakeAllComboBoxUpdate();
			// DrawPreview(); TO DO!
		}

		private void ConditionChangeHandler(object sender, PropertyChangedEventArgs e)
		{
			if (Rvt.Control.Condition == RvtControl.Status.Idle)
			{
				EnableControls(true);
				AvailableCategories.Items.Refresh();
				PickedCategories.Items.Refresh();
				AvailableElements.Items.Refresh();
				PickedElements.Items.Refresh();
			}
			else if (Rvt.Control.Condition == RvtControl.Status.Export)
				EnableControls(false);
			else if (Rvt.Control.Condition == RvtControl.Status.Error)
			{
				MessageBox.Show(String.Format("{0}\n\nStack trace:\n{1}",
					Rvt.Control.LastException.Message, Rvt.Control.LastException.StackTrace),
						"Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
			}
			else if (Rvt.Control.Condition == RvtControl.Status.Sort)
				EnableControls(false);
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
			Rvt.Control.Condition = RvtControl.Status.Export;

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
				Rvt.Control.Condition = RvtControl.Status.Error;
			}
		}

		private void ExportWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			ExportProgressBar.Value = e.ProgressPercentage;
		}

		private void ExportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			if (e.Error != null || Rvt.Control.Condition == RvtControl.Status.Error)
			{
				Rvt.Control.Condition = RvtControl.Status.Error;
			}
			else
			{
				Rvt.Control.SaveProcedure();    // Save file
				Rvt.Data.InitExportElements();	// Unset Rvt.Data.ExportElements
			}
			ExportProgressBar.Value = 0.0;
			Rvt.Control.Condition = RvtControl.Status.Idle;
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
			// DrawPreview(); TO DO!
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

		private void AvailableCategoriesNameColumnHeader_Click(object sender, RoutedEventArgs e)
		{
			Rvt.Data.AvailableCategories.Sort(true);
			//Rvt.Control.Condition = RvtControl.Status.Sort;
		}

		private void AvailableCategoriesCountColumnHeader_Click(object sender, RoutedEventArgs e)
		{
			Rvt.Data.AvailableCategories.Sort(false);
		}
	} // class MainWindow

} // namespace RevitToGOST
