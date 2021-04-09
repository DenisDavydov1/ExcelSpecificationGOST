using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace RevitToGOST
{
	public static partial class Work
	{
		public static Bitmaps Bitmaps { get; set; }
	}

	public class Bitmaps
	{
		/*
		** Member properties
		*/

		public BitmapImage Title { get; set; }
		public BitmapImage Table { get; set; }
		public BitmapImage Stamp { get; set; }
		public BitmapImage Dop { get; set; }

		public static readonly BitmapImage[] Previews =
		{
			new BitmapImage(new Uri("Previews/Empty.png", UriKind.Relative)),							// None,
			new BitmapImage(new Uri("Previews/Preview_GOST_21_110_2013_Table1.png", UriKind.Relative)),	// GOST_21_110_2013_Table1
			null,																						// GOST_P_21_101_2020_Dop3
			new BitmapImage(new Uri("Previews/Preview_GOST_21_101_2020_Stamp3.png", UriKind.Relative)),	// GOST_P_21_101_2020_Stamp3
			null																						// GOST_P_21_101_2020_Title_12
		};

		/*
		** Member methods
		*/

	} // class Bitmaps



	//static class PreviewImages
	//{
	//	public static readonly Bitmap[] Images =
	//	{
	//		Resources.Preview_Empty,
	//		Resources.Preview_GOST_21_110_2013_Table1,
	//		null,
	//		Resources.Preview_GOST_21_101_2020_Stamp3,
	//		null
	//	};
	//}

	//static partial class Work
	//{
	//	public static Previews Previews { get; set; }
	//}

	//public class Previews
	//{
	//	public Bitmap Title { get; set; }
	//	public Bitmap Table { get; set; }
	//	public Bitmap Stamp { get; set; }
	//	public Bitmap Dop { get; set; }

	//	public Previews()
	//	{
	//		Title = PreviewImages.Images[(int)GOST.Standarts.None];
	//		Table = PreviewImages.Images[(int)GOST.Standarts.GOST_21_110_2013_Table1];
	//		Stamp = PreviewImages.Images[(int)GOST.Standarts.None];
	//		Dop = PreviewImages.Images[(int)GOST.Standarts.None];
	//	}
	//}

} // namespace RevitToGOST
