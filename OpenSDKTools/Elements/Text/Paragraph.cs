using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class Paragraph : IElement, IWritable
	{
		public Paragraph()
		{
			this.Key           = string.Empty;
			this.Content       = new List<IText>();
			this.TextAlignment = TextAlignment.Left;
			this.IsNumbering   = false;
			this.NumberingId   = 0;
			this.HasHyperlink  = false;
			this.SpaceAfter    = ParagraphSpacing.None;
		}

		public bool IsNumbering { get; set; }
		internal bool HasHyperlink { get; set; }

		public string Key { get; set; }

		public ParagraphStyle Style { get; set; }

		public ParagraphSpacing SpaceAfter { get; set; }

		public TextAlignment TextAlignment { get; set; }

		internal List<IText> Content { get; set; }

		internal int NumberingId { get; set; }

		public void AddText(Text text)
		{
			this.Content.Add(text);
		}

		public void AddLink(Link link)
		{
			this.HasHyperlink = true;
			this.Content.Add(link);
		}

		public void AddBreak(BreakType? type = null)
		{
			this.Content.Add(new Break(type));
		}
		
		IWriter IWritable.GetWriter()
		{
			return new Word.ParagraphWriter();
		}
	}
}
