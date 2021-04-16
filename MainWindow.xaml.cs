using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
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
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitToGOST
{
	public partial class MainWindow : Window
	{
		#region properties

		private CollectionView PickedElementsView { get; set; }
		private CollectionView AvailableElementsView { get; set; }

		#endregion properties

		/////// MAIN WINDOW ///////
		#region main window methods

		public MainWindow()
		{
			InitializeComponent();

			Rvt.Windows.Condition = RvtWindows.Status.Idle;
			Rvt.Windows.ConditionChanged += new PropertyChangedEventHandler(ConditionChangeHandler);
			Work.Bitmaps.PreviewPageChanged += new PropertyChangedEventHandler(PreviewPageChangeHandler);
			Closing += WindowClosing;

			AvailableCategories.ItemsSource = Rvt.Data.AvailableCategories;
			PickedCategories.ItemsSource = Rvt.Data.PickedCategories;
			AvailableElements.ItemsSource = Rvt.Data.AvailableElements;
			PickedElements.ItemsSource = Rvt.Data.PickedElements;
			MakeAllComboBoxUpdate();

			DrawPreview();
		}

		/*
		** Closing window event handler
		*/

		private void WindowClosing(object sender, CancelEventArgs e)
		{
			if (Rvt.Control.ExportWorker != null)
			{
				Rvt.Control.ExportWorker.CancelAsync();
			}
		}

		/*
		** Application states handler
		*/

		private void ConditionChangeHandler(object sender, PropertyChangedEventArgs e)
		{
			if (Rvt.Windows.Condition == RvtWindows.Status.Idle)
			{
				ExportProgressBar.Value = 0.0;
				Export.Content = "Экспорт";
				EnableControls(true);
				AvailableCategories.Items.Refresh();
				PickedCategories.Items.Refresh();
				AvailableElements.Items.Refresh();
				PickedElements.Items.Refresh();
			}
			else if (Rvt.Windows.Condition == RvtWindows.Status.Export)
			{
				EnableControls(false);
				ExportProgressBar.Visibility = System.Windows.Visibility.Visible;
			}
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
			//Export.IsEnabled = state;
			GroupElemsCheckBox.IsEnabled = state;
			EnumerateColumnsCheckBox.IsEnabled = state;

			// Tab 1
			TitleComboBox.IsEnabled = state;
			TableComboBox.IsEnabled = state;
			Stamp1ComboBox.IsEnabled = state;
			Dop1ComboBox.IsEnabled = state;
			Stamp2ComboBox.IsEnabled = state;
			Dop2ComboBox.IsEnabled = state;

			// Tab 2
			AvailableCategories.IsEnabled = state;
			PickedCategories.IsEnabled = state;
			PickAllCategoriesButton.IsEnabled = state;
			ReleaseAllCategoriesButton.IsEnabled = state;

			// Tab 3
			AvailableElements.IsEnabled = state;
			PickedElements.IsEnabled = state;
			PickAllElementsButton.IsEnabled = state;
			ReleaseAllElementsButton.IsEnabled = state;
		}

		/*
		** Export button and worker methods
		*/

		private void Export_Click(object sender, RoutedEventArgs e)
		{
			if (Rvt.Windows.Condition == RvtWindows.Status.Idle)
			{
				Rvt.Windows.Condition = RvtWindows.Status.Export;

				BackgroundWorker exportWorker = new BackgroundWorker();
				exportWorker.WorkerReportsProgress = true;
				exportWorker.WorkerSupportsCancellation = true;
				exportWorker.DoWork += ExportWorker_DoWork;
				exportWorker.ProgressChanged += ExportWorker_ProgressChanged;
				exportWorker.RunWorkerCompleted += ExportWorker_RunWorkerCompleted;

				Export.Content = "Отмена";

				exportWorker.RunWorkerAsync();
			}
			else if (Rvt.Windows.Condition == RvtWindows.Status.Export)
			{
				if (Rvt.Control.ExportWorker != null)
				{
					Rvt.Control.ExportWorker.CancelAsync();
				}
			}
		}

		private void ExportWorker_DoWork(object sender, DoWorkEventArgs e)
		{
			try
			{
				Rvt.Control.ExportProcedure(sender, e);
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
				Rvt.Handler.Result = Result.Failed;
			}
			else if (e.Cancelled == true)
			{
				Rvt.Data.InitExportElements();  // Unset Rvt.Data.ExportElements
				Rvt.Handler.Result = Result.Cancelled;
			}
			else
			{
				Rvt.Control.SaveProcedure();    // Save file
				Rvt.Data.InitExportElements();  // Unset Rvt.Data.ExportElements
				Rvt.Handler.Result = Result.Succeeded;
			}
			Rvt.Windows.Condition = RvtWindows.Status.Idle;
		}

		/*
		** Группировать элементы (CheckBox)
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

		#endregion main window methods

		/////// TAB CONTROL ///////
		#region tab control

		private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Tab1 == null || Tab2 == null || Tab2 == null)
				return;
			Tab1.Header = RvtWindows.TabNames[0];
			Tab2.Header = RvtWindows.TabNames[1];
			Tab3.Header = RvtWindows.TabNames[2];
			Tab1.Height = 40;
			Tab2.Height = 40;
			Tab3.Height = 40;

			if (MainTabControl.SelectedIndex == 0)
			{
				Tab1.Header += RvtWindows.TabDescr[0];
				Tab1.Height = 160;
			}
			else if (MainTabControl.SelectedIndex == 1)
			{
				Tab2.Header += RvtWindows.TabDescr[1];
				Tab2.Height = 160;
			}
			else if (MainTabControl.SelectedIndex == 2)
			{
				Tab3.Header += RvtWindows.TabDescr[2];
				Tab3.Height = 160;
			}
		}

		#endregion tab control

		/////// TAB 1 ///////
		#region tab 1 methods

		/*
		** Настройка конфигурации таблицы (ComboBox)
		*/

		private void MakeAllComboBoxUpdate()
		{
			TitleComboBox_SelectionChanged(null, null);
			TableComboBox_SelectionChanged(null, null);
			Stamp1ComboBox_SelectionChanged(null, null);
			Dop1ComboBox_SelectionChanged(null, null);
			Stamp2ComboBox_SelectionChanged(null, null);
			Dop2ComboBox_SelectionChanged(null, null);
		}

		private void TitleComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (TitleComboBox.SelectedIndex == 0)
			{
				Work.Book.Title = GOST.Standarts.None;
				if (Work.Bitmaps.PreviewPage > 1)
					Work.Bitmaps.PreviewPage--;
			}
			else if (TitleComboBox.SelectedIndex == 1)
				Work.Book.Title = GOST.Standarts.GOST_P_21_101_2020_Title_12;
			DrawPreview();
		}

		private void TableComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (TableComboBox.SelectedIndex == 0)
				Work.Book.Table = GOST.Standarts.GOST_21_110_2013_Table_1;
			else if (TableComboBox.SelectedIndex == 1)
				Work.Book.Table = GOST.Standarts.GOST_P_21_101_2020_Table_7;
			else if (TableComboBox.SelectedIndex == 2)
				Work.Book.Table = GOST.Standarts.GOST_P_21_101_2020_Table_8;
			DrawPreview();
		}

		private void Stamp1ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Stamp1ComboBox.SelectedIndex == 0)
				Work.Book.Stamp1 = GOST.Standarts.None;
			else if (Stamp1ComboBox.SelectedIndex == 1)
				Work.Book.Stamp1 = GOST.Standarts.GOST_P_21_101_2020_Stamp_3;
			else if (Stamp1ComboBox.SelectedIndex == 2)
				Work.Book.Stamp1 = GOST.Standarts.GOST_P_21_101_2020_Stamp_4;
			else if (Stamp1ComboBox.SelectedIndex == 3)
				Work.Book.Stamp1 = GOST.Standarts.GOST_P_21_101_2020_Stamp_5;
			else if (Stamp1ComboBox.SelectedIndex == 4)
				Work.Book.Stamp1 = GOST.Standarts.GOST_P_21_101_2020_Stamp_6;
			else if (Stamp1ComboBox.SelectedIndex == 5)
				Work.Book.Stamp1 = GOST.Standarts.GOST_2_104_2006_Stamp_1;
			else if (Stamp1ComboBox.SelectedIndex == 6)
				Work.Book.Stamp1 = GOST.Standarts.GOST_2_104_2006_Stamp_2;
			else if (Stamp1ComboBox.SelectedIndex == 7)
				Work.Book.Stamp1 = GOST.Standarts.GOST_2_104_2006_Stamp_2a;
			DrawPreview();
		}

		private void Dop1ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Dop1ComboBox.SelectedIndex == 0)
				Work.Book.Dop1 = GOST.Standarts.None;
			else if (Dop1ComboBox.SelectedIndex == 1)
				Work.Book.Dop1 = GOST.Standarts.GOST_P_21_101_2020_Dop_3;
			else if (Dop1ComboBox.SelectedIndex == 2)
				Work.Book.Dop1 = GOST.Standarts.GOST_P_21_101_2020_Dop_4;
			else if (Dop1ComboBox.SelectedIndex == 3)
				Work.Book.Dop1 = GOST.Standarts.GOST_P_21_101_2020_Dop_5;
			else if (Dop1ComboBox.SelectedIndex == 4)
				Work.Book.Dop1 = GOST.Standarts.GOST_P_21_101_2020_Dop_6;
			else if (Dop1ComboBox.SelectedIndex == 5)
				Work.Book.Dop1 = GOST.Standarts.GOST_2_104_2006_Dop_1;
			else if (Dop1ComboBox.SelectedIndex == 6)
				Work.Book.Dop1 = GOST.Standarts.GOST_2_104_2006_Dop_2;
			else if (Dop1ComboBox.SelectedIndex == 7)
				Work.Book.Dop1 = GOST.Standarts.GOST_2_104_2006_Dop_2a;
			DrawPreview();
		}

		private void Stamp2ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Stamp2ComboBox.SelectedIndex == 0)
				Work.Book.Stamp2 = GOST.Standarts.None;
			else if (Stamp2ComboBox.SelectedIndex == 1)
				Work.Book.Stamp2 = GOST.Standarts.GOST_P_21_101_2020_Stamp_3;
			else if (Stamp2ComboBox.SelectedIndex == 2)
				Work.Book.Stamp2 = GOST.Standarts.GOST_P_21_101_2020_Stamp_4;
			else if (Stamp2ComboBox.SelectedIndex == 3)
				Work.Book.Stamp2 = GOST.Standarts.GOST_P_21_101_2020_Stamp_5;
			else if (Stamp2ComboBox.SelectedIndex == 4)
				Work.Book.Stamp2 = GOST.Standarts.GOST_P_21_101_2020_Stamp_6;
			else if (Stamp2ComboBox.SelectedIndex == 5)
				Work.Book.Stamp2 = GOST.Standarts.GOST_2_104_2006_Stamp_1;
			else if (Stamp2ComboBox.SelectedIndex == 6)
				Work.Book.Stamp2 = GOST.Standarts.GOST_2_104_2006_Stamp_2;
			else if (Stamp2ComboBox.SelectedIndex == 7)
				Work.Book.Stamp2 = GOST.Standarts.GOST_2_104_2006_Stamp_2a;
			DrawPreview();
		}

		private void Dop2ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (Dop2ComboBox.SelectedIndex == 0)
				Work.Book.Dop2 = GOST.Standarts.None;
			else if (Dop2ComboBox.SelectedIndex == 1)
				Work.Book.Dop2 = GOST.Standarts.GOST_P_21_101_2020_Dop_3;
			else if (Dop2ComboBox.SelectedIndex == 2)
				Work.Book.Dop2 = GOST.Standarts.GOST_P_21_101_2020_Dop_4;
			else if (Dop2ComboBox.SelectedIndex == 3)
				Work.Book.Dop2 = GOST.Standarts.GOST_P_21_101_2020_Dop_5;
			else if (Dop2ComboBox.SelectedIndex == 4)
				Work.Book.Dop2 = GOST.Standarts.GOST_P_21_101_2020_Dop_6;
			else if (Dop2ComboBox.SelectedIndex == 5)
				Work.Book.Dop2 = GOST.Standarts.GOST_2_104_2006_Dop_1;
			else if (Dop2ComboBox.SelectedIndex == 6)
				Work.Book.Dop2 = GOST.Standarts.GOST_2_104_2006_Dop_2;
			else if (Dop2ComboBox.SelectedIndex == 7)
				Work.Book.Dop2 = GOST.Standarts.GOST_2_104_2006_Dop_2a;
			DrawPreview();
		}

		/*
		** Preview controls and event 
		*/

		private void DrawPreview()
		{
			if (ImageTable == null || ImageStamp == null || ImageDop == null)
				return;
			PreviewPageNumberTextBox.Text = Work.Bitmaps.PreviewPageString();
			Work.Bitmaps.UpdateImages();
			ImageTable.Source = Work.Bitmaps.Table;
			ImageStamp.Source = Work.Bitmaps.Stamp;
			ImageDop.Source = Work.Bitmaps.Dop;
			UpdatePageDescriprion();
		}

		private void PreviewPageChangeHandler(object sender, PropertyChangedEventArgs e)
		{
			DrawPreview();
		}

		private void PrevPageButton_Click(object sender, RoutedEventArgs e)
		{
			int newPage = Work.Bitmaps.PreviewPage - 1;
			if (newPage < 1)
				return;
			Work.Bitmaps.PreviewPage = newPage;
		}

		private void NextPageButton_Click(object sender, RoutedEventArgs e)
		{
			int newPage = Work.Bitmaps.PreviewPage + 1;
			if (newPage > Work.Bitmaps.MaxPreviewPage)
				return;
			Work.Bitmaps.PreviewPage = newPage;
		}

		private void PreviewPageNumberTextBox_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Key == Key.Enter || e.Key == Key.Return)
			{
				int newPage = Work.Bitmaps.PreviewPage;
				try
				{
					newPage = Convert.ToInt32(PreviewPageNumberTextBox.Text);
				}
				catch { }
				if (newPage < 1)
					Work.Bitmaps.PreviewPage = 1;
				else if (newPage > Work.Bitmaps.MaxPreviewPage)
					Work.Bitmaps.PreviewPage = Work.Bitmaps.MaxPreviewPage;
				else
					Work.Bitmaps.PreviewPage = newPage;
			}
		}

		/*
		** Update page description
		*/

		private void UpdatePageDescriprion()
		{
			DescrTextBlock.Text = String.Empty;
			if (Work.Bitmaps.PreviewPage == 1)
			{
				if (Work.Book.Title != GOST.Standarts.None)
				{
					DescrTextBlock.Text += "Титульный лист:\n" + ConfFile.Descriprions[(int)Work.Book.Title];
				}
				else
				{
					DescrTextBlock.Text += "Спецификация:\n" + ConfFile.Descriprions[(int)Work.Book.Table] + "\n";
					if (Work.Book.Stamp1 != GOST.Standarts.None)
						DescrTextBlock.Text += "Основная надпись:\n" + ConfFile.Descriprions[(int)Work.Book.Stamp1] + "\n";
					if (Work.Book.Dop1 != GOST.Standarts.None)
						DescrTextBlock.Text += "Доп. графы:\n" + ConfFile.Descriprions[(int)Work.Book.Dop1] + "\n";
				}
			}
			else if (Work.Bitmaps.PreviewPage == 2)
			{
				if (Work.Book.Title != GOST.Standarts.None)
				{
					DescrTextBlock.Text += "Спецификация:\n" + ConfFile.Descriprions[(int)Work.Book.Table] + "\n";
					if (Work.Book.Stamp1 != GOST.Standarts.None)
						DescrTextBlock.Text += "Основная надпись:\n" + ConfFile.Descriprions[(int)Work.Book.Stamp1] + "\n";
					if (Work.Book.Dop1 != GOST.Standarts.None)
						DescrTextBlock.Text += "Доп. графы:\n" + ConfFile.Descriprions[(int)Work.Book.Dop1] + "\n";
				}
				else
				{
					DescrTextBlock.Text += "Спецификация:\n" + ConfFile.Descriprions[(int)Work.Book.Table] + "\n";
					if (Work.Book.Stamp2 != GOST.Standarts.None)
						DescrTextBlock.Text += "Основная надпись:\n" + ConfFile.Descriprions[(int)Work.Book.Stamp2] + "\n";
					if (Work.Book.Dop2 != GOST.Standarts.None)
						DescrTextBlock.Text += "Доп. графы:\n" + ConfFile.Descriprions[(int)Work.Book.Dop2] + "\n";
				}
			}
			else if (Work.Bitmaps.PreviewPage == 3)
			{
				DescrTextBlock.Text += "Спецификация:\n" + ConfFile.Descriprions[(int)Work.Book.Table] + "\n";
				if (Work.Book.Stamp2 != GOST.Standarts.None)
					DescrTextBlock.Text += "Основная надпись:\n" + ConfFile.Descriprions[(int)Work.Book.Stamp2] + "\n";
				if (Work.Book.Dop2 != GOST.Standarts.None)
					DescrTextBlock.Text += "Доп. графы:\n" + ConfFile.Descriprions[(int)Work.Book.Dop2] + "\n";
			}
		}

		#endregion tab 1 methods

		/////// TAB 2 ///////
		#region tab 2 methods

		/*
		**	Buttons Выбрать и освободить все категории
		*/

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

		#endregion tab 2 methods

		/////// TAB 3 ///////
		#region tab 3 methods

		/*
		**	Buttons Выбрать и освободить все элементы
		*/

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

		#endregion tab 3 methods

		/////// MISC ///////
		#region misc methods

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

		#endregion misc methods

		/////// HINT TRAY ///////
		#region hint tray

		private static T FindVisualParent<T>(UIElement element) where T : UIElement
		{
			UIElement parent = element;
			while (parent != null)
			{
				var correctlyTyped = parent as T;
				if (correctlyTyped != null)
				{
					return correctlyTyped;
				}
				parent = VisualTreeHelper.GetParent(parent) as UIElement;
			}
			return null;
		}

		private static T GetElementUnderMouse<T>() where T : UIElement
		{
			return FindVisualParent<T>(Mouse.DirectlyOver as UIElement);
		}

		private void Window_MouseMove(object sender, MouseEventArgs e)
		{
			if (GetElementUnderMouse<Button>() != null)
				SetHint_Button(GetElementUnderMouse<Button>());
			else if (GetElementUnderMouse<CheckBox>() != null)
				SetHint_CheckBox(GetElementUnderMouse<CheckBox>());
			else if (GetElementUnderMouse<ListView>() != null)
				SetHint_ListView(GetElementUnderMouse<ListView>());
			else if (GetElementUnderMouse<TabItem>() != null)
				SetHint_TabItem(GetElementUnderMouse<TabItem>());
			else if (GetElementUnderMouse<System.Windows.Controls.ComboBox>() != null)
				SetHint_ComboBox(GetElementUnderMouse<System.Windows.Controls.ComboBox>());
			else if (GetElementUnderMouse<TextBlock>() != null)
				SetHint_TextBlock(GetElementUnderMouse<TextBlock>());
			else if (GetElementUnderMouse<System.Windows.Controls.Image>() != null)
				SetHint_Image(GetElementUnderMouse<System.Windows.Controls.Image>());
			else if (GetElementUnderMouse<System.Windows.Controls.TextBox>() != null)
				SetHint_TextBox(GetElementUnderMouse<System.Windows.Controls.TextBox>());
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_Button(Button obj)
		{
			if (obj == Export && obj.Content as string == "Экспорт")
				HintTrayTextBlock.Text = "Начать экспорт выбранных даных в Excel";
			else if (obj == Export)
				HintTrayTextBlock.Text = "Отмена операции";
			else if (obj == PrevPageButton)
				HintTrayTextBlock.Text = "Показать предыдущий лист";
			else if (obj == NextPageButton)
				HintTrayTextBlock.Text = "Показать следующий лист";
			else if (obj == PickAllCategoriesButton)
				HintTrayTextBlock.Text = "Выбрать все доступные категории";
			else if (obj == ReleaseAllCategoriesButton)
				HintTrayTextBlock.Text = "Убрать все выбранные категории";
			else if (obj == PickAllElementsButton)
				HintTrayTextBlock.Text = "Выбрать все доступные элементы";
			else if (obj == ReleaseAllElementsButton)
				HintTrayTextBlock.Text = "Убрать все выбранные элементы";
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_CheckBox(CheckBox obj)
		{
			if (obj == EnumerateColumnsCheckBox)
				HintTrayTextBlock.Text = "Включить нумерацию столбцов на каждом листе спецификации";
			else if (obj == GroupElemsCheckBox)
				HintTrayTextBlock.Text = "Группировать элементы по категориям (в разделе выбора элементов и в спецификации)";
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_ListView(ListView obj)
		{
			if (obj == AvailableCategories)
				HintTrayTextBlock.Text = "Для выбора, перетащите категорию в список справа. Для множестенного выбора зажмите CTRL или SHIFT";
			else if (obj == PickedCategories)
				HintTrayTextBlock.Text = "Для исключения категории, перетащите ее в список слева";
			else if (obj == AvailableElements)
				HintTrayTextBlock.Text = "Для выбора, перетащите элемент в список справа. Для множестенного выбора зажмите CTRL или SHIFT";
			else if (obj == PickedElements)
				HintTrayTextBlock.Text = "Чтобы изменить порядок, перетащите элемент. Для исключения элемента, перетащите его в список слева";
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_TabItem(TabItem obj)
		{
			if (obj == Tab1)
				HintTrayTextBlock.Text = ""; // to do if i will not make tab description in window
			else if (obj == Tab2)
				HintTrayTextBlock.Text = "";
			else if (obj == Tab3)
				HintTrayTextBlock.Text = "";
			//else if (obj == Tab4)
			//	HintTrayTextBlock.Text = "";
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_ComboBox(System.Windows.Controls.ComboBox obj)
		{
			if (obj == TitleComboBox)
				HintTrayTextBlock.Text = "Стандарт титульного листа документации. Если титульный лист не требуется, выберите вариант <Нет>";
			else if (obj == TableComboBox)
				HintTrayTextBlock.Text = "Стандарт таблицы спецификации";
			else if (obj == Stamp1ComboBox)
				HintTrayTextBlock.Text = "Стандарт основной надписи для 1-го листа документации. Если не требуется, выберите вариант <Нет>";
			else if (obj == Dop1ComboBox)
				HintTrayTextBlock.Text = "Стандарт дополнительных граф для 1-го листа документации. Если не требуется, выберите вариант <Нет>";
			else if (obj == Stamp2ComboBox)
				HintTrayTextBlock.Text = "Стандарт основной надписи со 2-го листа документации. Если не требуется, выберите вариант <Нет>";
			else if (obj == Dop2ComboBox)
				HintTrayTextBlock.Text = "Стандарт дополнительных граф со 2-го листа документации. Если не требуется, выберите вариант <Нет>";
			else
				HintTrayTextBlock.Text = "";
		}

		private void SetHint_TextBlock(TextBlock obj)
		{
			HintTrayTextBlock.Text = "";
		}

		private void SetHint_Image(System.Windows.Controls.Image obj)
		{
			HintTrayTextBlock.Text = "";
		}

		private void SetHint_TextBox(System.Windows.Controls.TextBox obj)
		{
			if (obj == PreviewPageNumberTextBox)
				HintTrayTextBlock.Text = "Номер листа документации";
		}

		#endregion hint tray

	} // class MainWindow

} // namespace RevitToGOST
