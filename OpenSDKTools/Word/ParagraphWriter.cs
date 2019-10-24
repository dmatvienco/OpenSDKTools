using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenSDKTools.Word
{
	class ParagraphWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Paragraph;
		}

		public void Write(Marker marker, Paragraph paragraph)
		{
			Write(marker, paragraph, true);
		}

		public void Write(Marker marker, Paragraph paragraph, bool removeMarker)
		{
			if (paragraph == null)
			{
				return;
			}
			var markerParagraph = marker.Container;
			var parent = markerParagraph.Parent;

			parent.InsertAfter(
				GetParagraph(paragraph),
				markerParagraph
			);

			if (removeMarker)
				markerParagraph.Remove();
		}

		public void WriteTo(OpenXmlElement parent, Paragraph paragraph)
		{
			parent.Append(
				GetParagraph(paragraph)
			);
		}

		private W.Paragraph GetParagraph(Paragraph paragraph)
		{
			var p = new W.Paragraph();
			p.Append(new ParagraphProperties());

			if (paragraph.IsNumbering && paragraph.NumberingId > 0)
			{
				p.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };
				p.ParagraphProperties.NumberingProperties = new NumberingProperties();
				p.ParagraphProperties.NumberingProperties.Append(new NumberingLevelReference() { Val = 0 });
				p.ParagraphProperties.NumberingProperties.Append(new NumberingId() { Val = paragraph.NumberingId });
			}
			else
			{
				p.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = paragraph.Style.ToString() };
				p.ParagraphProperties.Justification = new Justification() { Val = GetJustificationValue(paragraph.TextAlignment) };
			}

			if(paragraph.SpaceAfter == ParagraphSpacing.Space10pt)
			{
				p.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();
				p.ParagraphProperties.SpacingBetweenLines.After = "200";
				p.ParagraphProperties.ContextualSpacing = new ContextualSpacing() { Val = false };
			}

			foreach (var text in paragraph.Content)
			{
				OpenXmlElement e = null;
				if (text is Text)
				{
					e = TextWriter.Generate(text as Text);
				}
				if (text is Link)
				{
					e = LinkWriter.Generate(text as Link);;
				}

				if (text is Break)
				{
					e = SymbolWriter.GenerateBreak(text as Break);
				}

				if (e != null)
					p.Append(e);
			}

			return p;
		}		

		private JustificationValues GetJustificationValue(TextAlignment alignment)
		{
			switch (alignment)
			{
				case TextAlignment.Center:
					return JustificationValues.Center;
				case TextAlignment.Right:
					return JustificationValues.Right;
				default:
					return JustificationValues.Left;
			}
		}
	}
}
