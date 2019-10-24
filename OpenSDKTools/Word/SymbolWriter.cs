using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Word
{
	class SymbolWriter
	{
		public static OpenXmlElement GenerateBreak(Break _break)
		{
			var b = new W.Break();
			if (_break.Type.HasValue)
			{
				switch(_break.Type.Value)
				{
					case BreakType.Column:
						b.Type = W.BreakValues.Column;
						break;
					case BreakType.Page:
						b.Type = W.BreakValues.Page;
						break;
					case BreakType.TextWrapping:
						b.Type = W.BreakValues.TextWrapping;
						break;
				}
			}

			return new W.Run(b);
		}
	}
}
