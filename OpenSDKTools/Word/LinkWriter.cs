using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Word
{
	class LinkWriter
	{
		public static OpenXmlElement Generate(Link link)
		{
			OpenXmlElement container = new Run();

			bool hasBookmark = false;

			if (!string.IsNullOrEmpty(link.Anchor))
			{
				hasBookmark = true;

				var b = new BookmarkStart() { Name = link.Anchor };
				b.Id = b.GetHashCode().ToString();
				var e = new BookmarkEnd() { Id = b.Id };

				container.Append(b, e);
			}

			if (!string.IsNullOrEmpty(link.Href))
			{
				var hyperlink = new Hyperlink();
				string href = string.Empty;

				if (!string.IsNullOrEmpty(link.HyperlinkRelationship))
				{
					hyperlink.Id = link.HyperlinkRelationship;
				}
				else
				{					
					hyperlink.Anchor = link.Href;
				}

				foreach (var item in link.Content)
				{
					var run = TextWriter.Generate(item);
					run.RunProperties.Append(
						new RunStyle()
						{
							Val = "Hyperlink"
						});

					hyperlink.Append(run);
				}

				if (!hasBookmark)
				{
					container = hyperlink;
				}
				else
				{
					container.Append(hyperlink);
				}
			}

			return container;
		}
	}
}
