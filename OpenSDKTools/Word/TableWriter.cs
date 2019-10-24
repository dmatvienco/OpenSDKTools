using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OpenSDKTools.Word
{
	class TableWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Table;
		}

		public void Write(Marker marker, Table table)
		{
			if (table == null)
				return;

			var markerParagraph = marker.Container;
			var parent = markerParagraph.Parent;

			var cellWriter = new TableCellWriter();

			var tableElement = new W.Table();

			var tableProp = new TableProperties();

			if (!string.IsNullOrEmpty(table.Width))
			{
				var tableStyle = new TableStyle() { Val = "Table" };

				// Make the table width 100% of the page width.
				var tableWidth = new TableWidth() { Width = table.Width, Type = TableWidthUnitValues.Pct };

				// Apply
				tableProp.Append(tableStyle, tableWidth);
				tableElement.AppendChild(tableProp);
			}

			AppendMargins(tableProp, table.TopMargin, table.LeftMargin, table.BottomMargin, table.RightMargin);

			foreach (var row in table.Rows)
			{
				var rowElement = new W.TableRow();

				if (row.Height > 0)
				{
					rowElement.Append( 
						new TableRowProperties(
							new TableRowHeight()
							{
								Val = (UInt32Value)(row.Height * 1U)
							}
						)
					);
				}

				foreach (var cell in row.Cells)
				{
					cellWriter.WriteTo(rowElement, cell);
				}

				tableElement.AppendChild(rowElement);
			}

			if(table.PageBreak == PageBreak.Before)
			{
				insertPageBreak(parent, markerParagraph);
			}

			parent.InsertAfter(tableElement, markerParagraph);
		}

		private static void insertPageBreak(OpenXmlElement parent, OpenXmlElement markerParagraph)
		{
			var pagebreak = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Break() { Type = BreakValues.Page }));
			parent.InsertAfter(pagebreak, markerParagraph);
		}

		private void AppendMargins(TableProperties tableProp, int topMargin, int leftMargin, int bottomMargin, int rightMargin)
		{
			var tableCellMarginDefault = new TableCellMarginDefault();
			var tableCellTopMargin = new TopMargin() { Width = topMargin.ToString(), Type = TableWidthUnitValues.Dxa };
			var tableCellLeftMargin = new TableCellLeftMargin() { Width = (Int16Value)leftMargin, Type = TableWidthValues.Dxa };
			var tableCellBottomMargin = new BottomMargin() { Width = bottomMargin.ToString(), Type = TableWidthUnitValues.Dxa };
			var tableCellRightMargin = new TableCellRightMargin() { Width = (Int16Value)rightMargin, Type = TableWidthValues.Dxa };

			tableCellMarginDefault.Append(tableCellTopMargin);
			tableCellMarginDefault.Append(tableCellLeftMargin);
			tableCellMarginDefault.Append(tableCellBottomMargin);
			tableCellMarginDefault.Append(tableCellRightMargin);

			tableProp.Append(tableCellMarginDefault);
		}
	}
}
