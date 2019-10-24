using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace OpenSDKTools.Word
{
    class DocumentWriter
	{
		internal void Write(string file, Dictionary<string, Container> containers, Dictionary<string, Variable> variables)
		{
			using (var package = WordprocessingDocument.Open(file, true))
			{
				var context = new Context();

				context.Package = package;
				context.Variables = variables;
				context.Containers = containers;

				//
				context.Markers = GetMarkers(context);

				if (context.Markers.Count > 0)
				{
					//
					HyperLink.Append(context);

					//
					FillMarkers(context);
				}
			}
		}

		private List<Marker> GetMarkers(Context context)
		{
			var mainPart = context.Package.MainDocumentPart;

			var markers = new List<SdtElement>();

			markers = mainPart.Document.Descendants<SdtElement>().ToList();

			// header
			if (mainPart.HeaderParts != null)
			{
				foreach (var part in mainPart.HeaderParts)
				{
					var _e = part.Header.Descendants<SdtElement>();

					markers.AddRange(_e);
				}
			}

			// footer
			if (mainPart.FooterParts != null)
			{
				foreach (var part in mainPart.FooterParts)
				{
					var _e = part.Footer.Descendants<SdtElement>();

					markers.AddRange(_e);
				}
			}

			// footnotes
			if (mainPart.FootnotesPart != null && mainPart.FootnotesPart.Footnotes != null)
			{
				foreach (var part in mainPart.FootnotesPart.Footnotes)
				{
					var _e = part.Descendants<SdtElement>();

					markers.AddRange(_e);
				}
			}

			return markers.ConvertAll<Marker>(x => { return Marker.Create(x); });
		}

		private void FillMarkers(Context context)
		{
			foreach (var marker in context.Markers)
			{
				if (marker.Tag == string.Empty)
					continue;

				Fill(context, marker);
			}
		}

		private void Fill(Context context, Marker marker)
		{
			var tag = marker.Element.SdtProperties.GetFirstChild<Tag>();
			if (tag == null)
			{
				return;
			}

			if (context.Variables.ContainsKey(tag.Val.Value))
			{
				var variable = context.Variables[tag.Val.Value];
				Fill(marker, variable);
				return;
			}

			if (context.Containers.ContainsKey(tag.Val.Value))
			{
				var container = context.Containers[tag.Val.Value];

				if (container.IsNumbering)
				{
					container.NumberingId = Numbering.Append(context);
				}

				foreach (var image in container.Children.OfType<OpenSDKTools.Image>())
				{
					image.Id = Image.Append(context, image.Data);
				}

				Fill(context, marker, container);
				return;
			}
		}

		private void Fill(Marker marker, Variable variable)
		{
			var writer = variable.GetWriter();

			switch (writer.GetWriterType())
			{
				case WriterType.Variable:
					(writer as VariableWriter).Write(marker, variable.Value);
					break;
				default:
					throw new NotImplementedException("This type is not supported : " + writer.GetWriterType());
					break;
			}

			marker.Element.Remove();
		}

		private void Fill(Context context, Marker marker, Container container)
		{
			int count = container.Children.Count - 1;
			for (int i = count; i > -1; i--)
			{
				var child = container.Children[i];

				var writer = child.GetWriter();

				switch (writer.GetWriterType())
				{
					case WriterType.Paragraph:
						var p = child as Paragraph;
						p.NumberingId = p.IsNumbering && container.IsNumbering && container.NumberingId > 0 ? container.NumberingId : 0;

						(writer as ParagraphWriter).Write(marker, p, false);
						break;
					case WriterType.Table:
						(writer as TableWriter).Write(marker, child as Table);
						break;
					case WriterType.Image:
						(writer as ImageWriter).Write(marker, child as OpenSDKTools.Image);
						break;
					case WriterType.Chart:
						(writer as ChartWriter).Write(context.Package, marker, child as OpenSDKTools.Chart);
						break;
					default:
						throw new NotImplementedException("This type is not supported : " + writer.GetWriterType());
						break;
				}
			}

			marker.Element.Remove();
		}

		private class Image
		{
			public static string Append(Context context, byte[] imageData)
			{
				string id = string.Format("imageId{0}", context.ImageId++);
				ImagePart imagePart = context.Package.MainDocumentPart.AddNewPart<ImagePart>("image/png", id);
				GenerateImagePartContent(imagePart, imageData);

				return id;
			}

			private static void GenerateImagePartContent(ImagePart imagePart, byte[] imageData)
			{
				System.IO.Stream data = new System.IO.MemoryStream(imageData);
				imagePart.FeedData(data);
				data.Close();
			}
		}

		private class Numbering
		{

			public static int Append(Context context)
			{
				var numbering = context.Package.MainDocumentPart.NumberingDefinitionsPart.Numbering;
				//
				AbstractNum anum = new AbstractNum() { AbstractNumberId = context.AbstructNumberingId };
				anum.MultiLevelType = new MultiLevelType() { Val = MultiLevelValues.Multilevel };

				var level = new Level();
				level.StartNumberingValue = new StartNumberingValue() { Val = 1 };
				level.LevelText = new LevelText() { Val = "%1." };
				level.LevelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };
				level.LevelIndex = 0;
				level.PreviousParagraphProperties = new PreviousParagraphProperties(new Indentation() { Left = "360" });

				anum.Append(level);

				var numberingId = context.NumberingId;

				var num = new NumberingInstance(new AbstractNumId() { Val = context.AbstructNumberingId }) { NumberID = numberingId };

				var last = numbering.Descendants<AbstractNum>().Last();

				numbering.InsertAfter(anum, last);
				numbering.Append(num);

				//
				context.NumberingId++;
				context.AbstructNumberingId++;

				return numberingId;
			}

		}

		private class HyperLink
		{
			public static void Append(Context context)
			{
				foreach (var pair in context.Containers)
				{
					foreach (var p in pair.Value.Children)
					{
						if (p is Paragraph)
						{
							Append(context, p as Paragraph);
						}
						if (p is Table)
						{
							Append(context, p as Table);
						}
					}
				}
			}

			private static void Append(Context context, Table table)
			{
				foreach (var row in table.Rows)
				{
					foreach (var cell in row.Cells)
					{
						Append(context, cell.Content);
					}
				}
			}

			private static void Append(Context context, Paragraph paragraph)
			{
				if (paragraph.HasHyperlink)
				{
					foreach (var el in paragraph.Content)
					{
						if (el is Link)
						{
							Append(context, (Link)el);
						}
					}
				}
			}

			private static void Append(Context context, Link link)
			{
				link.Anchor = link.Name.Replace(" ", "_").TrimStart('#');

				if (!string.IsNullOrEmpty(link.Href))
				{
					if (link.Href.StartsWith("#"))
					{
						link.Href = link.Href.Replace(" ", "_").TrimStart('#');
					}
					else
					{
						var rel = context.Package.MainDocumentPart.AddHyperlinkRelationship(new Uri(link.Href), true);
						link.HyperlinkRelationship = rel.Id;
					}
				}
			}
		}
	}
}
