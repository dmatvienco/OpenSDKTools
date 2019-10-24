using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools
{
	public class Link : IText
	{
		public Link()
		{
			this.Href = string.Empty;
			this.Name = string.Empty;
			this.Anchor = string.Empty;

			this.HyperlinkRelationship = string.Empty;
			this.Content = new List<Text>();
		}

		public string Href { get; set; }
		public string Name { get; set; }

		internal string HyperlinkRelationship { get; set; }
		internal string Anchor { get; set; }

		public List<Text> Content { get; set; }
	}
}
