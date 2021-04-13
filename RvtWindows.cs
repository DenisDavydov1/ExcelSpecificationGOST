using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitToGOST
{
	static partial class Rvt
	{
		public static RvtWindows Windows;
	}

	public class RvtWindows
	{
		#region properties

		public MainWindow MainWindow { get; set; }

		public LoadingWindow LoadingWindow { get; set; }

		private Status _Condition; // Application status

		public Status Condition
		{
			get { return _Condition; }
			set
			{
				if (value != _Condition)
				{
					_Condition = value;
					OnConditionChanged();
				}
			}
		}

		#endregion properties

		public enum Status
		{
			Loading,
			Idle,
			Export,
			Error,
			Sort
		}

		public static readonly string[] TabNames = {
			"1. Настройка документа",
			"2. Выбор категорий",
			"3. Выбор элементов"
		};

		public static readonly string[] TabDescr = {
			"\n\n    Определение\n    стандартов для\n    составления\n    документации",
			"\n\n    Выбор категорий\n    объектов модели,\n    которые требуется\n    включить в\n    спецификацию",
			"\n\n    Формирование\n    списка объектов\n    модели, которые\n    будут включены\n    в спецификацию"
		};

		#region events

		public event PropertyChangedEventHandler ConditionChanged;

		protected void OnConditionChanged(string propertyName = null)
		{
			ConditionChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		#endregion events

		#region methods

		public RvtWindows() { }

		public void RunLoadingWindow()
		{
			LoadingWindow = new LoadingWindow();
			Condition = Status.Loading;
			LoadingWindow.Show();
		}

		public void CloseLoadingWindow()
		{
			//while (Condition == Status.Loading)
			//	System.Threading.Thread.Sleep(500);
			LoadingWindow.Close();
		}

		public void RunMainWindow()
		{
			MainWindow = new MainWindow();
			MainWindow.ShowDialog();
		}

		#endregion events

	} // public class RvtWindows

} // namespace RevitToGOST
