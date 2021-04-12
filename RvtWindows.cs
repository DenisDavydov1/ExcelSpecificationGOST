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
		/*
		** Member properties
		*/

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

		public enum Status
		{
			Loading,
			Idle,
			Export,
			Error,
			Sort
		}

		/*
		** Member events
		*/

		public event PropertyChangedEventHandler ConditionChanged;

		protected void OnConditionChanged(string propertyName = null)
		{
			ConditionChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		/*
		** Member methods
		*/

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

	} // public class RvtWindows

} // namespace RevitToGOST
