using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using C15 = DocumentFormat.OpenXml.Office2013.Drawing.Chart;
using DocumentFormat.OpenXml.Spreadsheet;


namespace OpenSDKTools.Excel
{
	class ChartWriter : IWriter
	{
		public void Write(WorksheetPart part, Chart chart)
		{
			DrawingsPart drawingsPart1 = part.AddNewPart<DrawingsPart>("rId1");
			GenerateDrawingsPart1Content(drawingsPart1, chart);

			ChartPart chartPart1 = drawingsPart1.AddNewPart<ChartPart>("rId1");
			GenerateChartPart1Content(chartPart1, chart);

			var drawing = new Drawing() { Id = "rId1" };
			part.Worksheet.Append(drawing);
		}

		public WriterType GetWriterType()
		{
			return WriterType.Chart;
		}

		// Generates content of drawingsPart1.
		//private void GenerateDrawingsPartContent(DrawingsPart drawingsPart, Chart chart)
		//{
		//	var worksheetDrawing = new Xdr.WorksheetDrawing();
		//	worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		//	worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

		//	var twoCellAnchor = new Xdr.TwoCellAnchor();

		//	var fromMarker = new Xdr.FromMarker();
		//	var columnIdFrom = new Xdr.ColumnId();
		//	columnIdFrom.Text = chart.ColumnFrom.ToString();
		//	Xdr.ColumnOffset columnOffsetFrom = new Xdr.ColumnOffset();
		//	columnOffsetFrom.Text = "11908";
		//	var rowIdFrom = new Xdr.RowId();
		//	rowIdFrom.Text = chart.RowFrom.ToString();
		//	Xdr.RowOffset rowOffsetFrom = new Xdr.RowOffset();
		//	rowOffsetFrom.Text = "9523";

		//	fromMarker.Append(columnIdFrom);
		//	fromMarker.Append(columnOffsetFrom);
		//	fromMarker.Append(rowIdFrom);
		//	fromMarker.Append(rowOffsetFrom);

		//	var toMarker = new Xdr.ToMarker();
		//	var columnIdTo = new Xdr.ColumnId();
		//	columnIdTo.Text = chart.ColumnTo.ToString();
		//	Xdr.ColumnOffset columnOffsetTo = new Xdr.ColumnOffset();
		//	columnOffsetTo.Text = "250032";
		//	var rowIdTo = new Xdr.RowId();
		//	rowIdTo.Text = chart.RowTo.ToString();
		//	Xdr.RowOffset rowOffsetTo = new Xdr.RowOffset();
		//	rowOffsetTo.Text = "29764";

		//	toMarker.Append(columnIdTo);
		//	toMarker.Append(columnOffsetTo);
		//	toMarker.Append(rowIdTo);
		//	toMarker.Append(rowOffsetTo);

		//	var graphicFrame = new Xdr.GraphicFrame() { Macro = "" };

		//	var nonVisualGraphicFrameProperties = new Xdr.NonVisualGraphicFrameProperties();
		//	var nonVisualDrawingProperties = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = chart.TableKey };

		//	var nonVisualGraphicFrameDrawingProperties = new Xdr.NonVisualGraphicFrameDrawingProperties();
		//	var graphicFrameLocks = new A.GraphicFrameLocks();

		//	nonVisualGraphicFrameDrawingProperties.Append(graphicFrameLocks);

		//	nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
		//	nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);

		//	var transform = new Xdr.Transform();
		//	var offset = new A.Offset() { X = 0L, Y = 0L };
		//	var extents = new A.Extents() { Cx = 0L, Cy = 0L };

		//	transform.Append(offset);
		//	transform.Append(extents);

		//	var graphic = new A.Graphic();

		//	var graphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

		//	var chartReference = new C.ChartReference() { Id = "rId1" };
		//	chartReference.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		//	chartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

		//	graphicData.Append(chartReference);

		//	graphic.Append(graphicData);

		//	graphicFrame.Append(nonVisualGraphicFrameProperties);
		//	graphicFrame.Append(transform);
		//	graphicFrame.Append(graphic);
		//	var clientData = new Xdr.ClientData();

		//	twoCellAnchor.Append(fromMarker);
		//	twoCellAnchor.Append(toMarker);
		//	twoCellAnchor.Append(graphicFrame);
		//	twoCellAnchor.Append(clientData);

		//	worksheetDrawing.Append(twoCellAnchor);

		//	drawingsPart.WorksheetDrawing = worksheetDrawing;
		//}

		//// Generates content of chartPart1.
		//private void GenerateChartPartContent(ChartPart chartPart, Chart chart)
		//{
		//	var chartSpace = new C.ChartSpace();
		//	chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		//	chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
		//	chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		//	var editingLanguage = new C.EditingLanguage() { Val = "en-US" };
		//	var style = new C.Style() { Val = 10 };

		//	var _chart = new C.Chart();

		//	var title = new C.Title();

		//	var chartText = new C.ChartText();

		//	var richText = new C.RichText();
		//	var bodyProperties = new A.BodyProperties();
		//	var listStyle = new A.ListStyle();

		//	var paragraph = new A.Paragraph();

		//	var paragraphProperties = new A.ParagraphProperties();
		//	var defaultRunProperties = new A.DefaultRunProperties();

		//	paragraphProperties.Append(defaultRunProperties);

		//	var run = new A.Run();
		//	var runProperties = new A.RunProperties() { Language = "en-US", FontSize = 800 };
		//	var text = new A.Text();
		//	text.Text = chart.Title;

		//	run.Append(runProperties);
		//	run.Append(text);

		//	//A.Run run2 = new A.Run();
		//	//A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 800, Baseline = 0 };
		//	//A.Text text2 = new A.Text();
		//	//text2.Text = title;

		//	//run2.Append(runProperties2);
		//	//run2.Append(text2);
		//	var endParagraphRunProperties = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 800 };

		//	paragraph.Append(paragraphProperties);
		//	paragraph.Append(run);
		//	//paragraph1.Append(run2);
		//	paragraph.Append(endParagraphRunProperties);

		//	richText.Append(bodyProperties);
		//	richText.Append(listStyle);
		//	richText.Append(paragraph);

		//	chartText.Append(richText);

		//	var layout1 = new C.Layout();

		//	var manualLayout1 = new C.ManualLayout();
		//	var leftMode1 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
		//	var topMode1 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
		//	var left1 = new C.Left() { Val = 0.14936699324798144D };
		//	var top1 = new C.Top() { Val = 7.5867300613079197E-2D };


		//	manualLayout1.Append(leftMode1);
		//	manualLayout1.Append(topMode1);
		//	manualLayout1.Append(left1);
		//	manualLayout1.Append(top1);

		//	layout1.Append(manualLayout1);

		//	title.Append(chartText);
		//	title.Append(layout1);

		//	var plotArea = new C.PlotArea();

		//	var layout2 = new C.Layout();

		//	var manualLayout2 = new C.ManualLayout();
		//	var layoutTarget2 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
		//	var leftMode2 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
		//	var topMode2 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
		//	var left2 = new C.Left() { Val = 0.10245464404093282D };
		//	var top2 = new C.Top() { Val = 4.7416814491091287E-2D };
		//	var width2 = new C.Width() { Val = 0.88919609910728359D };

		//	// chart height inside word drawing part
		//	//C.Height height1 = new C.Height() { Val = 0.81899924741893582D }; // original generated value
		//	var height2 = new C.Height() { Val = 0.86 };

		//	manualLayout2.Append(layoutTarget2);
		//	manualLayout2.Append(leftMode2);
		//	manualLayout2.Append(topMode2);
		//	manualLayout2.Append(left2);
		//	manualLayout2.Append(top2);
		//	manualLayout2.Append(width2);
		//	manualLayout2.Append(height2);

		//	layout2.Append(manualLayout2);

		//	var areaChart = new C.AreaChart();
		//	var grouping = new C.Grouping() { Val = C.GroupingValues.Standard };

		//	var areaChartSeries = new C.AreaChartSeries();
		//	var index = new C.Index() { Val = (UInt32Value)0U };
		//	var order = new C.Order() { Val = (UInt32Value)0U };

		//	var seriesText = new C.SeriesText();

		//	var stringReference = new C.StringReference();
		//	var formula1 = new C.Formula();
		//	formula1.Text = chart.AxisX;

		//	var stringCache = new C.StringCache();
		//	var pointCount1 = new C.PointCount() { Val = (UInt32Value)1U };

		//	var stringPoint = new C.StringPoint() { Index = (UInt32Value)0U };
		//	var numericValue = new C.NumericValue();
		//	numericValue.Text = chart.LegendTitle;

		//	stringPoint.Append(numericValue);

		//	stringCache.Append(pointCount1);
		//	stringCache.Append(stringPoint);

		//	stringReference.Append(formula1);
		//	stringReference.Append(stringCache);

		//	seriesText.Append(stringReference);

		//	var values = new C.Values();

		//	var numberReference = new C.NumberReference();
		//	var formula2 = new C.Formula();
		//	formula2.Text = chart.AxisY;

		//	C.NumberingCache numberingCache = new C.NumberingCache();
		//	C.FormatCode formatCode = new C.FormatCode();
		//	formatCode.Text = "0.00%";

		//	/* years */
		//	C.PointCount pointCount2 = new C.PointCount() { Val = UInt32Value.FromUInt32((uint)chart.Labels.Count) };

		//	numberingCache.Append(formatCode);
		//	numberingCache.Append(pointCount2);

		//	for (int i = 0; i < chart.Labels.Count; i++)
		//	{
		//		C.NumericPoint numericPoint = new C.NumericPoint() { Index = UInt32Value.FromUInt32((uint)i) };
		//		C.NumericValue _numericValue = new C.NumericValue();
		//		numericValue.Text = string.Format("{0}E-2", chart.Labels[i]);
		//		numericPoint.Append(_numericValue);
		//		numberingCache.Append(numericPoint);
		//	}

		//	numberReference.Append(formula2);
		//	numberReference.Append(numberingCache);

		//	values.Append(numberReference);

		//	areaChartSeries.Append(index);
		//	areaChartSeries.Append(order);
		//	areaChartSeries.Append(seriesText);
		//	areaChartSeries.Append(values);
		//	var axisId1 = new C.AxisId() { Val = (UInt32Value)78173696U };
		//	var axisId2 = new C.AxisId() { Val = (UInt32Value)78175232U };

		//	areaChart.Append(grouping);
		//	areaChart.Append(areaChartSeries);
		//	areaChart.Append(axisId1);
		//	areaChart.Append(axisId2);

		//	var categoryAxis1 = new C.CategoryAxis();
		//	var axisId3 = new C.AxisId() { Val = (UInt32Value)78173696U };

		//	var scaling1 = new C.Scaling();
		//	var orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

		//	scaling1.Append(orientation1);
		//	var axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
		//	var majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
		//	var tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
		//	var crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)78175232U };
		//	var crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
		//	var autoLabeled1 = new C.AutoLabeled() { Val = true };
		//	var labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
		//	var labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };

		//	categoryAxis1.Append(axisId3);
		//	categoryAxis1.Append(scaling1);
		//	categoryAxis1.Append(axisPosition1);
		//	categoryAxis1.Append(majorTickMark1);
		//	categoryAxis1.Append(tickLabelPosition1);
		//	categoryAxis1.Append(crossingAxis1);
		//	categoryAxis1.Append(crosses1);
		//	categoryAxis1.Append(autoLabeled1);
		//	categoryAxis1.Append(labelAlignment1);
		//	categoryAxis1.Append(labelOffset1);

		//	var valueAxis1 = new C.ValueAxis();
		//	var axisId4 = new C.AxisId() { Val = (UInt32Value)78175232U };

		//	var scaling2 = new C.Scaling();
		//	var orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

		//	scaling2.Append(orientation2);
		//	var axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
		//	var majorGridlines1 = new C.MajorGridlines();
		//	var numberingFormat1 = new C.NumberingFormat() { FormatCode = "0.00%", SourceLinked = true };
		//	var majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
		//	var tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
		//	var crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)78173696U };
		//	var crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
		//	var crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };

		//	valueAxis1.Append(axisId4);
		//	valueAxis1.Append(scaling2);
		//	valueAxis1.Append(axisPosition2);
		//	valueAxis1.Append(majorGridlines1);
		//	valueAxis1.Append(numberingFormat1);
		//	valueAxis1.Append(majorTickMark2);
		//	valueAxis1.Append(tickLabelPosition2);
		//	valueAxis1.Append(crossingAxis2);
		//	valueAxis1.Append(crosses2);
		//	valueAxis1.Append(crossBetween1);

		//	var dataTable1 = new C.DataTable();
		//	var showHorizontalBorder1 = new C.ShowHorizontalBorder() { Val = true };
		//	var showVerticalBorder1 = new C.ShowVerticalBorder() { Val = true };
		//	var showOutlineBorder1 = new C.ShowOutlineBorder() { Val = true };
		//	var showKeys1 = new C.ShowKeys() { Val = true };

		//	dataTable1.Append(showHorizontalBorder1);
		//	dataTable1.Append(showVerticalBorder1);
		//	dataTable1.Append(showOutlineBorder1);
		//	dataTable1.Append(showKeys1);

		//	C.ShapeProperties shapeProperties1 = new C.ShapeProperties();

		//	A.Outline outline1 = new A.Outline();
		//	A.NoFill noFill1 = new A.NoFill();

		//	outline1.Append(noFill1);

		//	shapeProperties1.Append(outline1);

		//	plotArea.Append(layout2);
		//	plotArea.Append(areaChart);
		//	plotArea.Append(categoryAxis1);
		//	plotArea.Append(valueAxis1);
		//	plotArea.Append(dataTable1);
		//	plotArea.Append(shapeProperties1);
		//	var plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
		//	var displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Zero };

		//	_chart.Append(title);
		//	_chart.Append(plotArea);
		//	_chart.Append(plotVisibleOnly1);
		//	_chart.Append(displayBlanksAs1);

		//	var textProperties1 = new C.TextProperties();
		//	var bodyProperties2 = new A.BodyProperties();
		//	var listStyle2 = new A.ListStyle();

		//	var paragraph2 = new A.Paragraph();

		//	var paragraphProperties2 = new A.ParagraphProperties();
		//	var defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 700 };

		//	paragraphProperties2.Append(defaultRunProperties2);
		//	var endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

		//	paragraph2.Append(paragraphProperties2);
		//	paragraph2.Append(endParagraphRunProperties2);

		//	textProperties1.Append(bodyProperties2);
		//	textProperties1.Append(listStyle2);
		//	textProperties1.Append(paragraph2);

		//	var printSettings1 = new C.PrintSettings();
		//	var headerFooter1 = new C.HeaderFooter();
		//	var pageMargins1 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
		//	var pageSetup1 = new C.PageSetup() { Orientation = C.PageSetupOrientationValues.Landscape };

		//	printSettings1.Append(headerFooter1);
		//	printSettings1.Append(pageMargins1);
		//	printSettings1.Append(pageSetup1);

		//	chartSpace.Append(editingLanguage);
		//	chartSpace.Append(style);
		//	chartSpace.Append(_chart);
		//	chartSpace.Append(textProperties1);
		//	chartSpace.Append(printSettings1);

		//	var chartShapeProperties2 = new ChartShapeProperties();
		//	var outline2 = new DocumentFormat.OpenXml.Drawing.Outline();
		//	var noFill2 = new NoFill();
		//	outline2.Append(noFill2);
		//	chartShapeProperties2.Append(outline2);
		//	//chartSpace.Append(chartShapeProperties2);

		//	chartPart.ChartSpace = chartSpace;
		//}

		// Generates content of drawingsPart1.
		private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, Chart chart)
		{
			Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
			worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
			worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

			Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

			Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
			Xdr.ColumnId columnId1 = new Xdr.ColumnId();
			columnId1.Text = chart.ColumnFrom.ToString();
			Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
			columnOffset1.Text = "0";
			Xdr.RowId rowId1 = new Xdr.RowId();
			rowId1.Text = chart.RowFrom.ToString();
			Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
			rowOffset1.Text = "0";

			fromMarker1.Append(columnId1);
			fromMarker1.Append(columnOffset1);
			fromMarker1.Append(rowId1);
			fromMarker1.Append(rowOffset1);

			Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
			Xdr.ColumnId columnId2 = new Xdr.ColumnId();
			columnId2.Text = chart.ColumnTo.ToString();
			Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
			columnOffset2.Text = "238124";
			Xdr.RowId rowId2 = new Xdr.RowId();
			rowId2.Text = chart.RowTo.ToString();
			Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
			rowOffset2.Text = "20241";

			toMarker1.Append(columnId2);
			toMarker1.Append(columnOffset2);
			toMarker1.Append(rowId2);
			toMarker1.Append(rowOffset2);

			Xdr.GraphicFrame graphicFrame1 = new Xdr.GraphicFrame() { Macro = "" };

			Xdr.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new Xdr.NonVisualGraphicFrameProperties();

			Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = chart.TableKey };

			A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

			A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

			OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D4EA194D-E283-4B88-B3BE-83B3557FCE42}\" />");

			nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

			nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

			nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);

			Xdr.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Xdr.NonVisualGraphicFrameDrawingProperties();
			A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();

			nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

			nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties1);
			nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);

			Xdr.Transform transform1 = new Xdr.Transform();
			A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
			A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };

			transform1.Append(offset1);
			transform1.Append(extents1);

			A.Graphic graphic1 = new A.Graphic();

			A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

			C.ChartReference chartReference1 = new C.ChartReference() { Id = "rId1" };
			chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			chartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

			graphicData1.Append(chartReference1);

			graphic1.Append(graphicData1);

			graphicFrame1.Append(nonVisualGraphicFrameProperties1);
			graphicFrame1.Append(transform1);
			graphicFrame1.Append(graphic1);
			Xdr.ClientData clientData1 = new Xdr.ClientData();

			twoCellAnchor1.Append(fromMarker1);
			twoCellAnchor1.Append(toMarker1);
			twoCellAnchor1.Append(graphicFrame1);
			twoCellAnchor1.Append(clientData1);

			worksheetDrawing1.Append(twoCellAnchor1);

			drawingsPart1.WorksheetDrawing = worksheetDrawing1;
		}

		// Generates content of chartPart1.
		private void GenerateChartPart1Content(ChartPart chartPart1, Chart chart)
		{
			C.ChartSpace chartSpace1 = new C.ChartSpace();
			chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
			chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
			chartSpace1.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
			C.Date1904 date19041 = new C.Date1904() { Val = false };
			C.EditingLanguage editingLanguage1 = new C.EditingLanguage() { Val = "en-US" };
			C.RoundedCorners roundedCorners1 = new C.RoundedCorners() { Val = true };

			AlternateContent alternateContent1 = new AlternateContent();
			alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

			AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
			alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
			C14.Style style1 = new C14.Style() { Val = 110 };

			alternateContentChoice1.Append(style1);

			AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
			C.Style style2 = new C.Style() { Val = 10 };

			alternateContentFallback1.Append(style2);

			alternateContent1.Append(alternateContentChoice1);
			alternateContent1.Append(alternateContentFallback1);

			C.Chart chart1 = new C.Chart();

			C.Title title1 = new C.Title();

			C.ChartText chartText1 = new C.ChartText();

			C.RichText richText1 = new C.RichText();
			A.BodyProperties bodyProperties1 = new A.BodyProperties();
			A.ListStyle listStyle1 = new A.ListStyle();

			A.Paragraph paragraph1 = new A.Paragraph();

			A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();
			A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties();

			paragraphProperties1.Append(defaultRunProperties1);

			A.Run run1 = new A.Run();
			A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 800 };
			A.Text text1 = new A.Text();
			text1.Text = "";

			run1.Append(runProperties1);
			run1.Append(text1);

			A.Run run2 = new A.Run();
			A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 800, Baseline = 0 };
			A.Text text2 = new A.Text();
			text2.Text = chart.Title;

			run2.Append(runProperties2);
			run2.Append(text2);
			A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 800 };

			paragraph1.Append(paragraphProperties1);
			paragraph1.Append(run1);
			paragraph1.Append(run2);
			paragraph1.Append(endParagraphRunProperties1);

			richText1.Append(bodyProperties1);
			richText1.Append(listStyle1);
			richText1.Append(paragraph1);

			chartText1.Append(richText1);

			C.Layout layout1 = new C.Layout();

			C.ManualLayout manualLayout1 = new C.ManualLayout();
			C.LeftMode leftMode1 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
			C.TopMode topMode1 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
			C.Left left1 = new C.Left() { Val = 0.14936699324798144D };
			C.Top top1 = new C.Top() { Val = 7.5867300613079197E-2D };

			manualLayout1.Append(leftMode1);
			manualLayout1.Append(topMode1);
			manualLayout1.Append(left1);
			manualLayout1.Append(top1);

			layout1.Append(manualLayout1);
			C.Overlay overlay1 = new C.Overlay() { Val = true };

			title1.Append(chartText1);
			title1.Append(layout1);
			title1.Append(overlay1);
			C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

			C.PlotArea plotArea1 = new C.PlotArea();

			C.Layout layout2 = new C.Layout();

			C.ManualLayout manualLayout2 = new C.ManualLayout();
			C.LayoutTarget layoutTarget1 = new C.LayoutTarget() { Val = C.LayoutTargetValues.Inner };
			C.LeftMode leftMode2 = new C.LeftMode() { Val = C.LayoutModeValues.Edge };
			C.TopMode topMode2 = new C.TopMode() { Val = C.LayoutModeValues.Edge };
			C.Left left2 = new C.Left() { Val = 0.10245464404093282D };
			C.Top top2 = new C.Top() { Val = 4.7416814491091287E-2D };
			C.Width width1 = new C.Width() { Val = 0.88919609910728359D };
			C.Height height1 = new C.Height() { Val = 0.86D };

			manualLayout2.Append(layoutTarget1);
			manualLayout2.Append(leftMode2);
			manualLayout2.Append(topMode2);
			manualLayout2.Append(left2);
			manualLayout2.Append(top2);
			manualLayout2.Append(width1);
			manualLayout2.Append(height1);

			layout2.Append(manualLayout2);

			C.AreaChart areaChart1 = new C.AreaChart();
			C.Grouping grouping1 = new C.Grouping() { Val = C.GroupingValues.Standard };
			C.VaryColors varyColors1 = new C.VaryColors() { Val = true };

			C.AreaChartSeries areaChartSeries1 = new C.AreaChartSeries();
			C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
			C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

			C.SeriesText seriesText1 = new C.SeriesText();
			C.NumericValue numericValue1 = new C.NumericValue();
			numericValue1.Text = chart.LegendTitle;

			seriesText1.Append(numericValue1);

			C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

			C.NumberReference numberReference1 = new C.NumberReference();

			C.NumRefExtensionList numRefExtensionList1 = new C.NumRefExtensionList();

			C.NumRefExtension numRefExtension1 = new C.NumRefExtension() { Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" };
			numRefExtension1.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");

			C15.FullReference fullReference1 = new C15.FullReference();
			C15.SequenceOfReferences sequenceOfReferences1 = new C15.SequenceOfReferences();
			sequenceOfReferences1.Text = chart.AxisX;

			fullReference1.Append(sequenceOfReferences1);

			numRefExtension1.Append(fullReference1);

			numRefExtensionList1.Append(numRefExtension1);
			C.Formula formula1 = new C.Formula();
			formula1.Text = chart.AxisX;

			C.NumberingCache numberingCache1 = new C.NumberingCache();
			C.FormatCode formatCode1 = new C.FormatCode();
			formatCode1.Text = "General";
			C.PointCount pointCount1 = new C.PointCount() { Val = (UInt32Value)(uint)chart.Labels.Count };

			numberingCache1.Append(formatCode1);
			numberingCache1.Append(pointCount1);

			for (uint i = 0; i < chart.Labels.Count; i++)
			{
				C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)i };
				C.NumericValue numericValue2 = new C.NumericValue();
				numericValue2.Text = chart.Labels[(int)i];

				numericPoint1.Append(numericValue2);

				numberingCache1.Append(numericPoint1);
			}

			numberReference1.Append(numRefExtensionList1);
			numberReference1.Append(formula1);
			numberReference1.Append(numberingCache1);

			categoryAxisData1.Append(numberReference1);

			C.Values values1 = new C.Values();

			C.NumberReference numberReference2 = new C.NumberReference();

			C.NumRefExtensionList numRefExtensionList2 = new C.NumRefExtensionList();

			C.NumRefExtension numRefExtension2 = new C.NumRefExtension() { Uri = "{02D57815-91ED-43cb-92C2-25804820EDAC}" };
			numRefExtension2.AddNamespaceDeclaration("c15", "http://schemas.microsoft.com/office/drawing/2012/chart");

			C15.FullReference fullReference2 = new C15.FullReference();
			C15.SequenceOfReferences sequenceOfReferences2 = new C15.SequenceOfReferences();
			sequenceOfReferences2.Text = chart.AxisY;

			fullReference2.Append(sequenceOfReferences2);

			numRefExtension2.Append(fullReference2);

			numRefExtensionList2.Append(numRefExtension2);
			C.Formula formula2 = new C.Formula();
			formula2.Text = chart.AxisY;

			C.NumberingCache numberingCache2 = new C.NumberingCache();
			C.FormatCode formatCode2 = new C.FormatCode();
			formatCode2.Text = "0.00%";
			C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)(uint)chart.Values.Count };

			numberingCache2.Append(formatCode2);
			numberingCache2.Append(pointCount2);

			for (uint i = 0; i < chart.Values.Count; i++)
			{
				C.NumericPoint numericPoint27 = new C.NumericPoint() { Index = (UInt32Value)i };
				C.NumericValue numericValue28 = new C.NumericValue();
				numericValue28.Text = chart.Values[(int)i];

				numericPoint27.Append(numericValue28);

				numberingCache2.Append(numericPoint27);
			}
			
			numberReference2.Append(numRefExtensionList2);
			numberReference2.Append(formula2);
			numberReference2.Append(numberingCache2);

			values1.Append(numberReference2);

			C.AreaSerExtensionList areaSerExtensionList1 = new C.AreaSerExtensionList();

			C.AreaSerExtension areaSerExtension1 = new C.AreaSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
			areaSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

			OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c16:uniqueId val=\"{00000000-69FF-4CCD-9302-CEC5CC8046DF}\" xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" />");

			areaSerExtension1.Append(openXmlUnknownElement2);

			areaSerExtensionList1.Append(areaSerExtension1);

			areaChartSeries1.Append(index1);
			areaChartSeries1.Append(order1);
			areaChartSeries1.Append(seriesText1);
			areaChartSeries1.Append(categoryAxisData1);
			areaChartSeries1.Append(values1);
			areaChartSeries1.Append(areaSerExtensionList1);

			C.DataLabels dataLabels1 = new C.DataLabels();
			C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
			C.ShowValue showValue1 = new C.ShowValue() { Val = false };
			C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
			C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
			C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
			C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

			dataLabels1.Append(showLegendKey1);
			dataLabels1.Append(showValue1);
			dataLabels1.Append(showCategoryName1);
			dataLabels1.Append(showSeriesName1);
			dataLabels1.Append(showPercent1);
			dataLabels1.Append(showBubbleSize1);
			C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)78173696U };
			C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)78175232U };

			areaChart1.Append(grouping1);
			areaChart1.Append(varyColors1);
			areaChart1.Append(areaChartSeries1);
			areaChart1.Append(dataLabels1);
			areaChart1.Append(axisId1);
			areaChart1.Append(axisId2);

			C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
			C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)78173696U };

			C.Scaling scaling1 = new C.Scaling();
			C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

			scaling1.Append(orientation1);
			C.Delete delete1 = new C.Delete() { Val = true };
			C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };
			C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
			C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
			C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.Cross };
			C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
			C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)78175232U };
			C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
			C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
			C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
			C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
			C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = true };

			categoryAxis1.Append(axisId3);
			categoryAxis1.Append(scaling1);
			categoryAxis1.Append(delete1);
			categoryAxis1.Append(axisPosition1);
			categoryAxis1.Append(numberingFormat1);
			categoryAxis1.Append(majorTickMark1);
			categoryAxis1.Append(minorTickMark1);
			categoryAxis1.Append(tickLabelPosition1);
			categoryAxis1.Append(crossingAxis1);
			categoryAxis1.Append(crosses1);
			categoryAxis1.Append(autoLabeled1);
			categoryAxis1.Append(labelAlignment1);
			categoryAxis1.Append(labelOffset1);
			categoryAxis1.Append(noMultiLevelLabels1);

			C.ValueAxis valueAxis1 = new C.ValueAxis();
			C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)78175232U };

			C.Scaling scaling2 = new C.Scaling();
			C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

			scaling2.Append(orientation2);
			C.Delete delete2 = new C.Delete() { Val = true };
			C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
			C.MajorGridlines majorGridlines1 = new C.MajorGridlines();
			C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
			C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
			C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.Cross };
			C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };
			C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)78173696U };
			C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
			C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.MidpointCategory };

			valueAxis1.Append(axisId4);
			valueAxis1.Append(scaling2);
			valueAxis1.Append(delete2);
			valueAxis1.Append(axisPosition2);
			valueAxis1.Append(majorGridlines1);
			valueAxis1.Append(numberingFormat2);
			valueAxis1.Append(majorTickMark2);
			valueAxis1.Append(minorTickMark2);
			valueAxis1.Append(tickLabelPosition2);
			valueAxis1.Append(crossingAxis2);
			valueAxis1.Append(crosses2);
			valueAxis1.Append(crossBetween1);

			C.DataTable dataTable1 = new C.DataTable();
			C.ShowHorizontalBorder showHorizontalBorder1 = new C.ShowHorizontalBorder() { Val = true };
			C.ShowVerticalBorder showVerticalBorder1 = new C.ShowVerticalBorder() { Val = true };
			C.ShowOutlineBorder showOutlineBorder1 = new C.ShowOutlineBorder() { Val = true };
			C.ShowKeys showKeys1 = new C.ShowKeys() { Val = true };

			dataTable1.Append(showHorizontalBorder1);
			dataTable1.Append(showVerticalBorder1);
			dataTable1.Append(showOutlineBorder1);
			dataTable1.Append(showKeys1);

			C.ShapeProperties shapeProperties1 = new C.ShapeProperties();

			A.Outline outline1 = new A.Outline();
			A.NoFill noFill1 = new A.NoFill();

			outline1.Append(noFill1);

			shapeProperties1.Append(outline1);

			plotArea1.Append(layout2);
			plotArea1.Append(areaChart1);
			plotArea1.Append(categoryAxis1);
			plotArea1.Append(valueAxis1);
			plotArea1.Append(dataTable1);
			plotArea1.Append(shapeProperties1);
			C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
			C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Zero };
			C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = true };

			chart1.Append(title1);
			chart1.Append(autoTitleDeleted1);
			chart1.Append(plotArea1);
			chart1.Append(plotVisibleOnly1);
			chart1.Append(displayBlanksAs1);
			chart1.Append(showDataLabelsOverMaximum1);

			C.ShapeProperties shapeProperties2 = new C.ShapeProperties();

			A.Outline outline2 = new A.Outline();
			A.NoFill noFill2 = new A.NoFill();

			outline2.Append(noFill2);

			shapeProperties2.Append(outline2);

			C.TextProperties textProperties1 = new C.TextProperties();
			A.BodyProperties bodyProperties2 = new A.BodyProperties();
			A.ListStyle listStyle2 = new A.ListStyle();

			A.Paragraph paragraph2 = new A.Paragraph();

			A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();
			A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 700 };

			paragraphProperties2.Append(defaultRunProperties2);
			A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

			paragraph2.Append(paragraphProperties2);
			paragraph2.Append(endParagraphRunProperties2);

			textProperties1.Append(bodyProperties2);
			textProperties1.Append(listStyle2);
			textProperties1.Append(paragraph2);

			C.PrintSettings printSettings1 = new C.PrintSettings();
			C.HeaderFooter headerFooter1 = new C.HeaderFooter();
			C.PageMargins pageMargins1 = new C.PageMargins() { Left = 0.70000000000000018D, Right = 0.70000000000000018D, Top = 0.75000000000000022D, Bottom = 0.75000000000000022D, Header = 0.3000000000000001D, Footer = 0.3000000000000001D };
			C.PageSetup pageSetup1 = new C.PageSetup() { Orientation = C.PageSetupOrientationValues.Landscape };

			printSettings1.Append(headerFooter1);
			printSettings1.Append(pageMargins1);
			printSettings1.Append(pageSetup1);

			chartSpace1.Append(date19041);
			chartSpace1.Append(editingLanguage1);
			chartSpace1.Append(roundedCorners1);
			chartSpace1.Append(alternateContent1);
			chartSpace1.Append(chart1);
			chartSpace1.Append(shapeProperties2);
			chartSpace1.Append(textProperties1);
			chartSpace1.Append(printSettings1);

			chartPart1.ChartSpace = chartSpace1;
		}
	}
}
