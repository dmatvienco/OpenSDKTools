using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class TableRow
	{
		public TableRow()
		{
			Cells = new List<TableCell>();
		}

		public List<TableCell> Cells { get; set; }

		/// <summary>
		/// in 20th of a point (DXA)
		/// 0 - default
		/// </summary>
		public int Height { get; set; }
	}
}
