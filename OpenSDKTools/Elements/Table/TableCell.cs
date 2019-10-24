using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class TableCell
	{
		public TableCell()
		{
			this.Border = BorderStyle.None;
			this.TextVerticalAlignment = TextVerticalAlignment.Center;
			

		}

		public void SetParagraph(Paragraph paragraph)
		{
			this.Content = paragraph;
		}

		/// <summary>
		/// in 20th of a point (DXA)
		/// 0 - default
		/// </summary>
		public int Width { get; set; }

		internal Paragraph Content { get; set; }

		public BorderStyle Border { get; set; }

		public TextVerticalAlignment TextVerticalAlignment  { get; set; }

		public Settings.CellMargin CellMargin { get; set; } = new Settings.CellMargin();

		public int ColSpan { get; set; }

		public string Value { get; set; }

		public string Fill { get; set; }

		public bool NoWrap { get; set; } = false;
	}
}
