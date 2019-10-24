using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenSDKTools.Word
{
	class VariableWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Variable;
		}

		public void Write(Marker marker, string text)
		{
			var parent = marker.Container;

			var r = marker.Element.Descendants<Run>().First();
			var txt = r.GetFirstChild<W.Text>();
			txt.Text = text;

			parent.InsertAfter(r.CloneNode(true), marker.Element);
		}
	}
}
