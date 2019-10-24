using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Excel
{
	public class Variable
	{
		public uint RowIndex { get; set; }
		public string CellReference { get; set; }
		public string Value { get; set; }
		public string Key
		{
			get
			{
				return $"{RowIndex}:{CellReference}";
			}
		}
	}
}
