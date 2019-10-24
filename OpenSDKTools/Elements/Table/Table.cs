using System.Collections.Generic;

namespace OpenSDKTools
{
    public class Table : IElement, IWritable
	{
		public Table()
		{
			this.Key       = string.Empty;
			this.Rows      = new List<TableRow>();
			this.Width     = "5000";
			this.PageBreak = PageBreak.None;
		}

		public string Key { get; set; }

		public List<TableRow> Rows { get; set; }

		public string Width { get; set; }

		public int TopMargin = 58;
		public int LeftMargin = 101;
		public int BottomMargin = 14;
		public int RightMargin = 101;

		public PageBreak PageBreak;

		IWriter IWritable.GetWriter()
		{
			return new Word.TableWriter();
		}
	}
}
