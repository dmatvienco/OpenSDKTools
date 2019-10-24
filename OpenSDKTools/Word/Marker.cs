using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OpenSDKTools.Word
{
	class Marker
	{
		public OpenXmlElement Container { get; set; }
		public W.SdtElement Element { get; set; }
		public string Tag { get; set; }

		public static Marker Create(W.SdtElement element)
		{
			var ret = new Marker();

			ret.Tag = GetTag(element);
			ret.Element = element;
			//ret.Container = element.Parent;
			ret.Container = GetParagraph(element);

			return ret;
		}

		private static OpenXmlElement GetParagraph(OpenXmlElement marker)
		{
			if (marker == null)
			{
				return null;
			}

			if (marker is W.Paragraph)
			{
				return marker;
			}

			if (marker is W.SdtBlock)
			{
				return marker;
			}

			return GetParagraph(marker.Parent);
		}

		private static string GetTag(W.SdtElement marker)
		{
			var tag = marker.SdtProperties.GetFirstChild<W.Tag>();
			if (tag == null)
			{
				return string.Empty;
			}

			return tag.Val.Value;
		}
	}
}
