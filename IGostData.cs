using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitToGOST
{
	public interface IGostData
	{
		/*
		**	Member fields
		*/

		/*
		**	Member methods
		*/

		void FillLines();
		List<List<string>> FillList();
	} // public interface IGostData
} // namespace RevitToGOST
