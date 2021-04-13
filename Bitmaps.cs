using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;

namespace RevitToGOST
{
	public static partial class Work
	{
		public static Bitmaps Bitmaps { get; set; }
	}

	public class Bitmaps
	{
		#region properties

		public BitmapImage Table { get; set; }
		public BitmapImage Stamp { get; set; }
		public BitmapImage Dop { get; set; }

		private int _PreviewPage = 1;
		public int PreviewPage
		{
			get { return _PreviewPage; }
			set
			{
				//if (value != _PreviewPage)
				//{
					_PreviewPage = value;
					OnPreviewPageChanged();
				//}

			}
		}

		public int MaxPreviewPage
		{
			get
			{
				if (Work.Book.Title != GOST.Standarts.None)
					return 3;
				return 2;
			}
		}

		#endregion properties

		public static readonly BitmapImage[] Previews =
		{
			new BitmapImage(new Uri("Previews/Preview_Empty.png", UriKind.Relative)),						// None
			new BitmapImage(new Uri("Previews/Preview_GOST_21_110_2013_Table1.png", UriKind.Relative)),		// GOST_21_110_2013_Table1
			new BitmapImage(new Uri("Previews/Preview_GOST_21_101_2020_Stamp3.png", UriKind.Relative)),		// GOST_P_21_101_2020_Stamp3
			null,	// GOST_P_21_101_2020_Dop3
			null,	// stamp 4
			null,	// dop 4
			null,	// stamp 5
			null,	// dop 5
			null,	// stamp 6
			null,	// dop 6
			null,	// table 7
			null,	// table 8
			null,	// misc 9
			null,	// misc 9a
			null,	// misc 10
			null,	// misc 11
			new BitmapImage(new Uri("Previews/Preview_GOST_21_101_2020_Title12.png", UriKind.Relative)),	// GOST_P_21_101_2020_Title_12
			null,	// misc 13
			null	// title 14
		};

		#region events

		public event PropertyChangedEventHandler PreviewPageChanged;

		protected void OnPreviewPageChanged(string propertyName = null)
		{
			PreviewPageChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		#endregion events

		#region methods

		public void UpdateImages()
		{
			GOST.Standarts table = GOST.Standarts.None, stamp = GOST.Standarts.None, dop = GOST.Standarts.None;

			if (PreviewPage == 1)
			{
				if (Work.Book.Title != GOST.Standarts.None)
					table = Work.Book.Title;
				else
				{
					table = Work.Book.Table;
					stamp = Work.Book.Stamp1;
					dop = Work.Book.Dop1;
				}
			}
			else if (PreviewPage == 2)
			{
				if (Work.Book.Title != GOST.Standarts.None)
				{
					table = Work.Book.Table;
					stamp = Work.Book.Stamp1;
					dop = Work.Book.Dop1;
				}
				else
				{
					table = Work.Book.Table;
					stamp = Work.Book.Stamp2;
					dop = Work.Book.Dop2;
				}
			}
			else if (PreviewPage == 3)
			{
				if (Work.Book.Title != GOST.Standarts.None)
				{
					table = Work.Book.Table;
					stamp = Work.Book.Stamp2;
					dop = Work.Book.Dop2;
				}
			}
			Table = Previews[(int)table] ?? Previews[0];
			Stamp = Previews[(int)stamp] ?? Previews[0];
			Dop = Previews[(int)dop] ?? Previews[0];
		}

		public string PreviewPageString()
		{
			if (PreviewPage == MaxPreviewPage)
				return PreviewPage.ToString() + "...";
			return PreviewPage.ToString();
		}

		#endregion methods

	} // class Bitmaps

} // namespace RevitToGOST
