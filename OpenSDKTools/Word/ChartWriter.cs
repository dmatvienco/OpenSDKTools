using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using OVML = DocumentFormat.OpenXml.Vml.Office;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using SPRD = DocumentFormat.OpenXml.Spreadsheet;

namespace OpenSDKTools.Word
{
	class ChartWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Chart;
		}

		public void Write(WordprocessingDocument document, Marker marker, Chart chart)
		{
			var mainPart = document.MainDocumentPart;
			var parent = document.MainDocumentPart.Document.Body;

			var p = marker.Container;
			Run r = new Run();
			p.Append(r);
			Drawing drawing = new Drawing();
			r.Append(drawing);
			#region Import Chart
			DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline =
			new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
			new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent()
						//{ Cx = 6560000, Cy = 2880320 } //- live version
						//{ Cx = 6000000, Cy = 4500000 } //- last version
						{ Cx = (Int64Value)(6000000 * 1.4), Cy = (Int64Value)(3600000 * 1.4) } //- last version
							);
			byte[] byteArray = System.IO.File.ReadAllBytes(chart.ExcelFile);

			using (System.IO.MemoryStream mem = new System.IO.MemoryStream())
			{
				mem.Write(byteArray, 0, (int)byteArray.Length);

				//Open Excel spreadsheet
				using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(mem, true))
				{
					//Get all the appropriate parts
					WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;

					//WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
					WorksheetPart worksheetPart = Excel.DocumentWriter.GetWorksheetPart(mySpreadsheet, chart.SheetName);

					DrawingsPart drawingPart = worksheetPart.DrawingsPart;
					//ChartPart chartPart = drawingPart.ChartParts.First();
					ChartPart chartPart = drawingPart.ChartParts.First();

					//Clone the chart part and add it to my Word document
					ChartPart importedChartPart = mainPart.AddPart<ChartPart>(chartPart);
					string relId = mainPart.GetIdOfPart(importedChartPart);

					//chartPart.ChartSpace.ChildElements.First<DocumentFormat.OpenXml.Drawing.Charts.Chart>().Append(style);

					DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame frame =
						drawingPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().First();

					string chartName = frame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name;

					//Clone this node so we can add it to my Word document
					var clonedGraphic = (DocumentFormat.OpenXml.Drawing.Graphic)frame.Graphic.CloneNode(true);

					DocumentFormat.OpenXml.Drawing.Charts.ChartReference c =
						clonedGraphic.GraphicData.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>();
					c.Id = relId;

					//Give the chart a unique id and name
					var docPr = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties();
					docPr.Name = chartName;
					docPr.Id = Service.GetMaxDocPrId(mainPart) + 1;

					//add the chart data to the inline drawing object
					inline.Append(docPr, clonedGraphic);

					drawing.Append(inline);
				}
			}
			#endregion

			//parent.InsertBefore(p, marker.Element);
			//marker.Element.Remove();
		}
	}
}
