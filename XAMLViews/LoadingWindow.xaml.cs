using System.Threading;
using System.Windows;

namespace ExcelSpecificationGOST
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
		}
    }
}
