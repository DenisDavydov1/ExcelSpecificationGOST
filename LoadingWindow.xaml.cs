using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace RevitToGOST
{
	public partial class LoadingWindow : Window
	{
        public CancellationTokenSource LoadCancelToken { get; set; }

        public LoadingWindow()
		{
			InitializeComponent();
			LoadingScreenImage.Source = Bitmaps.Convert(Properties.Resources.loading_screen_image);
		}

		public new void Show()
		{
			base.Show();
			//System.Threading.Thread.Sleep(3000);
			//Rvt.Windows.Condition = RvtWindows.Status.Idle;
		}

    } // public partial class LoadingWindow
}
