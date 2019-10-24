using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class Chart : IWritable
	{
		public string TableKey { get; set; }
		public string SheetName { get; set; }
		public int AxisXRowNumber { get; set; }
		public int AxisYRowNumber { get; set; }
		public string LegendTitle { get; set; }
		public string Title { get; set; }
		public int RowFrom { get; set; }
		public int ColumnFrom { get; set; }
		public int RowTo { get; set; }
		public int ColumnTo { get; set; }
		public List<string> Labels { get; set; }
		public List<string> Values { get; set; }
		public string ExcelFile { get; set; }

		internal string AxisX { get; set; }
		internal string AxisY { get; set; }

		public IWriter GetWriter()
		{
			return new Word.ChartWriter();
		}
	}
}
