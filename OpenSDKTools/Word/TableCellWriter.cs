using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenSDKTools.Word
{
	class TableCellWriter
	{
		private ParagraphWriter paragraphWriter;

		public TableCellWriter()
		{
			this.paragraphWriter = new ParagraphWriter();
		}

		public void WriteTo(OpenXmlElement parent, TableCell cell)
		{
			var cellElement = new W.TableCell();
			cellElement.TableCellProperties = new TableCellProperties();

			ChangeMargin(cellElement, cell.CellMargin);

			AppendNoWrap(cellElement, cell.NoWrap);

			//
			AppendShading(cellElement, cell.Fill);

			//	
			AppendBorder(cellElement, cell.Border);

			//
			AppentWidth(cellElement, cell.Width);

			//
			AppendAlignment(cellElement, cell.TextVerticalAlignment);

			//
			AppentColSpan(cellElement, cell.ColSpan);

			//
			paragraphWriter.WriteTo(cellElement, cell.Content);

			//
			parent.Append(cellElement);
		}

		private void ChangeMargin(W.TableCell cell, Settings.CellMargin margin)		
		{
		

			if(!margin.MarginSet)
			{
				return;
			}

			var tableCellMargin = new TableCellMargin();

			if (margin.BottomMarginSet)
			{
				tableCellMargin.BottomMargin = new BottomMargin() { Width = margin.Bottom.ToString() };
			}

			if(margin.TopMarginSet)
			{
				tableCellMargin.TopMargin = new TopMargin() { Width = margin.Top.ToString() };
			}
			if(margin.LeftMarginSet)
			{ 
				tableCellMargin.LeftMargin = new LeftMargin() { Width = margin.Left.ToString() };
			}

			if(margin.RightMarginSet)
			{
				tableCellMargin.RightMargin = new RightMargin() { Width = margin.Right.ToString() };
			}

			cell.TableCellProperties.Append(tableCellMargin);

		}

		private static void AppendNoWrap(W.TableCell cell, bool noWrap)
		{
			if(noWrap)
			{
				var noWrapItem = new NoWrap();
				cell.TableCellProperties.Append(noWrapItem);
			}
		}

		private void AppendShading(W.TableCell cell, string fill)
		{
			if (!string.IsNullOrEmpty(fill))
			{
				var shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill };
				cell.TableCellProperties.Append(shading);
			}
		}

		private void AppendBorder(W.TableCell cell, BorderStyle style)
		{
			if((style & BorderStyle.Left) == BorderStyle.Left)
			{
				cell.TableCellProperties.Append(
					new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
				);
			}

			if ((style & BorderStyle.Top) == BorderStyle.Top)
			{
				cell.TableCellProperties.Append(
					new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
				);
			}

			if ((style & BorderStyle.Bottom) == BorderStyle.Bottom)
			{
				cell.TableCellProperties.Append(
					new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
				);
			}

			if ((style & BorderStyle.Right) == BorderStyle.Right)
			{
				cell.TableCellProperties.Append(
					new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U }
				);
			}
		}

		private void AppendAlignment(W.TableCell cell, TextVerticalAlignment alignment)
		{
			TableVerticalAlignmentValues al = TableVerticalAlignmentValues.Center;
			switch (alignment)
			{
				case TextVerticalAlignment.Top:
					al = TableVerticalAlignmentValues.Top;
					break;
				case TextVerticalAlignment.Center:
					al = TableVerticalAlignmentValues.Center;
					break;
				case TextVerticalAlignment.Bottom:
					al = TableVerticalAlignmentValues.Bottom;
					break;

			}

			cell.TableCellProperties.Append(
				new W.TableCellVerticalAlignment()
				{
					Val = al
				}
			);
		}

		private void AppentWidth(W.TableCell cell, int width)
		{
			if (width > 0)
			{
				cell.TableCellProperties.Append(
					new TableCellWidth()
					{
						Width = width.ToString(),
						Type = TableWidthUnitValues.Dxa
					}
				);
			}
		}

		private void AppentColSpan(W.TableCell cell, int colspan)
		{
			if (colspan > 1)
			{
				cell.TableCellProperties.Append(
					new GridSpan()
					{
						Val = colspan
					}
				);
			}
		}
	}
}
