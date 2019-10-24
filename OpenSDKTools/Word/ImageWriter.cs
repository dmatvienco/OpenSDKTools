using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenSDKTools.Word
{
	class ImageWriter : IWriter
	{
		public WriterType GetWriterType()
		{
			return WriterType.Image;
		}

		public void Write(Marker marker, Image image)
		{
			if (image == null)
				return;

			var markerParagraph = marker.Container;
			var parent = markerParagraph.Parent;

			W.Picture imageElement = new W.Picture();

			V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t75", CoordinateSize = "21600,21600", Filled = false, Stroked = false, OptionalNumber = 75, PreferRelative = true, EdgePath = "m@4@5l@4@11@9@11@9@5xe" };
			V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };

			V.Formulas formulas1 = new V.Formulas();
			V.Formula formula1 = new V.Formula() { Equation = "if lineDrawn pixelLineWidth 0" };
			V.Formula formula2 = new V.Formula() { Equation = "sum @0 1 0" };
			V.Formula formula3 = new V.Formula() { Equation = "sum 0 0 @1" };
			V.Formula formula4 = new V.Formula() { Equation = "prod @2 1 2" };
			V.Formula formula5 = new V.Formula() { Equation = "prod @3 21600 pixelWidth" };
			V.Formula formula6 = new V.Formula() { Equation = "prod @3 21600 pixelHeight" };
			V.Formula formula7 = new V.Formula() { Equation = "sum @0 0 1" };
			V.Formula formula8 = new V.Formula() { Equation = "prod @6 1 2" };
			V.Formula formula9 = new V.Formula() { Equation = "prod @7 21600 pixelWidth" };
			V.Formula formula10 = new V.Formula() { Equation = "sum @8 21600 0" };
			V.Formula formula11 = new V.Formula() { Equation = "prod @7 21600 pixelHeight" };
			V.Formula formula12 = new V.Formula() { Equation = "sum @10 21600 0" };

			formulas1.Append(formula1);
			formulas1.Append(formula2);
			formulas1.Append(formula3);
			formulas1.Append(formula4);
			formulas1.Append(formula5);
			formulas1.Append(formula6);
			formulas1.Append(formula7);
			formulas1.Append(formula8);
			formulas1.Append(formula9);
			formulas1.Append(formula10);
			formulas1.Append(formula11);
			formulas1.Append(formula12);
			V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle, AllowExtrusion = false };
			Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, AspectRatio = true };

			shapetype1.Append(stroke1);
			shapetype1.Append(formulas1);
			shapetype1.Append(path1);
			shapetype1.Append(lock1);

			V.Shape shape1 = new V.Shape() { Id = "_x0000_i1025", Style = string.Format("width:{0}pt;height:{1}pt", image.Width, image.Height), Type = "#_x0000_t75" };
			V.ImageData imageData1 = new V.ImageData() { Title = "eed80b48-ca3b-4af3-85c1-c5d1a0af4c63", RelationshipId = image.Id };

			shape1.Append(imageData1);

			imageElement.Append(shapetype1);
			imageElement.Append(shape1);

			parent.InsertAfter(imageElement, markerParagraph);
		}
	}
}
