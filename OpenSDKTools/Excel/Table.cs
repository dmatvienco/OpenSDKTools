using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Excel
{
	public class Table : OpenSDKTools.Table
	{
		public uint RowIndex { get; set; }
		public string ColumnName { get; set; }
	}
}
